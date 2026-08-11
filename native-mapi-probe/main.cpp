#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif

#include <windows.h>
#include <objidl.h>
#include <mapix.h>
#include <mapiutil.h>
#include <mapitags.h>

#include <fcntl.h>
#include <io.h>

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <cwctype>
#include <filesystem>
#include <iomanip>
#include <iostream>
#include <limits>
#include <memory>
#include <optional>
#include <sstream>
#include <stdexcept>
#include <string>
#include <unordered_set>
#include <utility>
#include <vector>

#pragma comment(lib, "MAPI32.lib")
#pragma comment(lib, "Ole32.lib")

namespace {

#if defined(_MSC_VER) && _MSC_VER < 1914
namespace fs = std::experimental::filesystem;
#else
namespace fs = std::filesystem;
#endif

static_assert(sizeof(void*) == 8, "NativeMapiProbe must be compiled as x64.");

constexpr std::size_t kQueryBatchSize = 64;
constexpr std::size_t kMaximumCliLimit = 1'000'000;
constexpr std::size_t kMaximumBodyChars = 4'000'000;
constexpr std::uint64_t kMaximumCliByteLimit = 1ULL << 40;  // 1 TiB hard ceiling.
constexpr std::uint64_t kDefaultMaxAttachmentBytes = 64ULL * 1024 * 1024;
constexpr std::uint64_t kDefaultMaxTotalAttachmentBytes = 0;  // Unlimited across streamed messages.
constexpr DWORD kAttachmentReadBufferBytes = 64 * 1024;

struct Options {
    std::size_t maxStores = 0;       // 0 means unlimited.
    std::size_t maxFolders = 100;
    std::size_t maxMessages = 20;
    std::size_t bodyPreviewChars = 240;
    bool jsonl = false;
    std::wstring storeContains;
    fs::path spoolDirectory;
    std::uint64_t maxAttachmentBytes = kDefaultMaxAttachmentBytes;
    std::uint64_t maxTotalAttachmentBytes = kDefaultMaxTotalAttachmentBytes;
};

struct Counters {
    std::size_t stores = 0;
    std::size_t folders = 0;
    std::size_t messages = 0;
    std::size_t skippedNonMail = 0;
    std::size_t errors = 0;
    std::size_t attachments = 0;
    std::size_t attachmentsSaved = 0;
    std::size_t attachmentsSkipped = 0;
    std::uint64_t attachmentBytesSaved = 0;
};

struct WinHandleCloser {
    void operator()(void* value) const noexcept {
        if (value != nullptr && value != INVALID_HANDLE_VALUE) {
            CloseHandle(value);
        }
    }
};

using WinHandle = std::unique_ptr<void, WinHandleCloser>;

template <typename T>
struct ComReleaser {
    void operator()(T* value) const noexcept {
        if (value != nullptr) {
            value->Release();
        }
    }
};

template <typename T>
using ComPtr = std::unique_ptr<T, ComReleaser<T>>;

struct MapiBufferReleaser {
    void operator()(void* value) const noexcept {
        if (value != nullptr) {
            MAPIFreeBuffer(value);
        }
    }
};

template <typename T>
using MapiBufferPtr = std::unique_ptr<T, MapiBufferReleaser>;

struct RowSetReleaser {
    void operator()(SRowSet* rows) const noexcept {
        if (rows != nullptr) {
            FreeProws(rows);
        }
    }
};

using RowSetPtr = std::unique_ptr<SRowSet, RowSetReleaser>;

class MapiRuntime final {
public:
    MapiRuntime() {
        MAPIINIT_0 init = {MAPI_INIT_VERSION, 0};
        const HRESULT result = MAPIInitialize(&init);
        if (FAILED(result)) {
            throw std::runtime_error("MAPIInitialize failed");
        }
        initialized_ = true;
    }

    MapiRuntime(const MapiRuntime&) = delete;
    MapiRuntime& operator=(const MapiRuntime&) = delete;

    ~MapiRuntime() {
        if (initialized_) {
            MAPIUninitialize();
        }
    }

private:
    bool initialized_ = false;
};

std::wstring HResultText(const HRESULT result) {
    std::wostringstream stream;
    stream << L"0x" << std::uppercase << std::hex << std::setw(8)
           << std::setfill(L'0') << static_cast<std::uint32_t>(result);
    return stream.str();
}

bool IsWithinLimit(const std::size_t value, const std::size_t limit) {
    return limit == 0 || value < limit;
}

bool ContainsCaseInsensitive(const std::wstring& value, const std::wstring& fragment) {
    if (fragment.empty()) {
        return true;
    }
    return std::search(
        value.begin(),
        value.end(),
        fragment.begin(),
        fragment.end(),
        [](const wchar_t left, const wchar_t right) {
            return std::towlower(left) == std::towlower(right);
        }) != value.end();
}

bool IsEmailMessageClass(const std::wstring& messageClass) {
    constexpr wchar_t kMailClass[] = L"IPM.Note";
    constexpr std::size_t kMailClassLength = (sizeof(kMailClass) / sizeof(wchar_t)) - 1;
    if (messageClass.size() < kMailClassLength) {
        return false;
    }
    for (std::size_t index = 0; index < kMailClassLength; ++index) {
        if (std::towlower(messageClass[index]) != std::towlower(kMailClass[index])) {
            return false;
        }
    }

    // Custom mail forms and S/MIME variants are IPM.Note.<suffix>. Requiring
    // the dot prevents unrelated classes such as IPM.Notebook from matching.
    // REPORT.* and IPM.Schedule.Meeting.* intentionally do not match: Outlook
    // materializes those as ReportItem/MeetingItem, while the previous OOM
    // indexer accepted only `item as Outlook.MailItem`.
    return messageClass.size() == kMailClassLength ||
        messageClass[kMailClassLength] == L'.';
}

std::wstring FromAnsi(const char* text, const std::size_t length = std::string::npos) {
    if (text == nullptr) {
        return {};
    }

    const std::size_t actualLength = length == std::string::npos
        ? std::char_traits<char>::length(text)
        : length;
    if (actualLength == 0) {
        return {};
    }
    if (actualLength > static_cast<std::size_t>(std::numeric_limits<int>::max())) {
        return L"<ANSI value too large>";
    }

    const int required = MultiByteToWideChar(
        CP_ACP,
        0,
        text,
        static_cast<int>(actualLength),
        nullptr,
        0);
    if (required <= 0) {
        return L"<ANSI conversion failed>";
    }

    std::wstring converted(static_cast<std::size_t>(required), L'\0');
    MultiByteToWideChar(
        CP_ACP,
        0,
        text,
        static_cast<int>(actualLength),
        converted.data(),
        required);
    return converted;
}

const SPropValue* FindProperty(
    const SPropValue* properties,
    const ULONG count,
    const ULONG propertyId,
    const ULONG preferredType = PT_UNSPECIFIED) {
    if (properties == nullptr) {
        return nullptr;
    }

    const SPropValue* fallback = nullptr;
    for (ULONG index = 0; index < count; ++index) {
        const SPropValue& property = properties[index];
        if (PROP_ID(property.ulPropTag) != propertyId ||
            PROP_TYPE(property.ulPropTag) == PT_ERROR) {
            continue;
        }
        if (preferredType == PT_UNSPECIFIED || PROP_TYPE(property.ulPropTag) == preferredType) {
            return &property;
        }
        fallback = &property;
    }
    return fallback;
}

std::wstring PropertyString(const SPropValue* property) {
    if (property == nullptr) {
        return {};
    }
    switch (PROP_TYPE(property->ulPropTag)) {
        case PT_UNICODE:
            return property->Value.lpszW == nullptr ? L"" : property->Value.lpszW;
        case PT_STRING8:
            return FromAnsi(property->Value.lpszA);
        default:
            return {};
    }
}

const SBinary* PropertyBinary(const SPropValue* property) {
    if (property == nullptr || PROP_TYPE(property->ulPropTag) != PT_BINARY) {
        return nullptr;
    }
    return &property->Value.bin;
}

std::wstring BinaryToHex(const SBinary* binary) {
    if (binary == nullptr || binary->cb == 0 || binary->lpb == nullptr) {
        return L"<missing>";
    }

    std::wostringstream stream;
    stream << std::uppercase << std::hex << std::setfill(L'0');
    for (ULONG index = 0; index < binary->cb; ++index) {
        stream << std::setw(2) << static_cast<unsigned int>(binary->lpb[index]);
    }
    return stream.str();
}

std::vector<BYTE> CopyBinary(const SBinary* binary) {
    if (binary == nullptr || binary->cb == 0 || binary->lpb == nullptr) {
        return {};
    }
    return {binary->lpb, binary->lpb + binary->cb};
}

std::vector<BYTE> ReadBinaryProperty(IMAPIProp* object, const ULONG propertyTag) {
    SizedSPropTagArray(1, tags) = {1, {propertyTag}};
    ULONG propertyCount = 0;
    LPSPropValue rawProperties = nullptr;
    const HRESULT status = object->GetProps(
        reinterpret_cast<LPSPropTagArray>(&tags),
        0,
        &propertyCount,
        &rawProperties);
    MapiBufferPtr<SPropValue> properties(rawProperties);
    if (FAILED(status) || rawProperties == nullptr) {
        return {};
    }
    return CopyBinary(PropertyBinary(FindProperty(
        rawProperties,
        propertyCount,
        PROP_ID(propertyTag),
        PT_BINARY)));
}

std::wstring CompactForConsole(std::wstring text) {
    for (wchar_t& character : text) {
        if (character == L'\r' || character == L'\n' || character == L'\t') {
            character = L' ';
        } else if (character < L' ' && character != L'\0') {
            character = L' ';
        }
    }
    text.erase(std::remove(text.begin(), text.end(), L'\0'), text.end());
    return text;
}

std::wstring JsonEscape(const std::wstring& text) {
    std::wostringstream escaped;
    escaped << std::hex << std::uppercase << std::setfill(L'0');
    for (const wchar_t character : text) {
        switch (character) {
            case L'"': escaped << L"\\\""; break;
            case L'\\': escaped << L"\\\\"; break;
            case L'\b': escaped << L"\\b"; break;
            case L'\f': escaped << L"\\f"; break;
            case L'\n': escaped << L"\\n"; break;
            case L'\r': escaped << L"\\r"; break;
            case L'\t': escaped << L"\\t"; break;
            default:
                // Escaping UTF-16 surrogate code units also makes malformed provider
                // strings harmless and preserves valid surrogate pairs in JSON.
                if (character < L' ' || (character >= 0xD800 && character <= 0xDFFF)) {
                    escaped << L"\\u" << std::setw(4)
                            << static_cast<unsigned int>(character);
                } else {
                    escaped << character;
                }
                break;
        }
    }
    return escaped.str();
}

std::wstring PropertyStringById(
    const SPropValue* properties,
    const ULONG count,
    const ULONG propertyId) {
    const SPropValue* property = FindProperty(properties, count, propertyId, PT_UNICODE);
    if (property == nullptr) {
        property = FindProperty(properties, count, propertyId, PT_STRING8);
    }
    return PropertyString(property);
}

std::wstring BinaryPropertyHex(
    const SPropValue* properties,
    const ULONG count,
    const ULONG propertyId) {
    const SBinary* binary = PropertyBinary(FindProperty(
        properties,
        count,
        propertyId,
        PT_BINARY));
    if (binary == nullptr || binary->cb == 0 || binary->lpb == nullptr) {
        return {};
    }
    return BinaryToHex(binary);
}

std::optional<std::wstring> PropertyUtcIso8601(
    const SPropValue* properties,
    const ULONG count,
    const ULONG propertyId) {
    const SPropValue* property = FindProperty(properties, count, propertyId, PT_SYSTIME);
    if (property == nullptr) {
        return std::nullopt;
    }

    SYSTEMTIME utc = {};
    if (!FileTimeToSystemTime(&property->Value.ft, &utc)) {
        return std::nullopt;
    }

    wchar_t buffer[32] = {};
    const int written = swprintf_s(
        buffer,
        L"%04hu-%02hu-%02huT%02hu:%02hu:%02hu.%03huZ",
        utc.wYear,
        utc.wMonth,
        utc.wDay,
        utc.wHour,
        utc.wMinute,
        utc.wSecond,
        utc.wMilliseconds);
    if (written <= 0) {
        return std::nullopt;
    }
    return std::wstring(buffer);
}

void PrintJsonNullableString(const std::optional<std::wstring>& value) {
    if (value.has_value()) {
        std::wcout << L'"' << JsonEscape(*value) << L'"';
    } else {
        std::wcout << L"null";
    }
}

std::wstring SanitizeAttachmentName(const std::wstring& rawName) {
    std::wstring name = rawName;
    const std::size_t separator = name.find_last_of(L"/\\");
    if (separator != std::wstring::npos) {
        name.erase(0, separator + 1);
    }

    for (wchar_t& character : name) {
        if (character < L' ' ||
            character == L'<' || character == L'>' || character == L':' ||
            character == L'"' || character == L'/' || character == L'\\' ||
            character == L'|' || character == L'?' || character == L'*') {
            character = L'_';
        }
    }
    while (!name.empty() && (name.back() == L'.' || name.back() == L' ')) {
        name.pop_back();
    }
    if (name.empty() || name == L"." || name == L"..") {
        name = L"attachment.bin";
    }
    constexpr std::size_t kMaximumSafeNameChars = 96;
    if (name.size() > kMaximumSafeNameChars) {
        name.resize(kMaximumSafeNameChars);
        if (!name.empty() && name.back() >= 0xD800 && name.back() <= 0xDBFF) {
            name.pop_back();
        }
        while (!name.empty() && (name.back() == L'.' || name.back() == L' ')) {
            name.pop_back();
        }
    }
    return name.empty() ? L"attachment.bin" : name;
}

fs::path PrepareSpoolDirectory(const fs::path& requested) {
    if (requested.empty()) {
        return {};
    }

    std::error_code error;
    fs::create_directories(requested, error);
    if (error || !fs::is_directory(requested, error) || error) {
        throw std::invalid_argument("cannot create or access --spool-dir");
    }

    fs::path canonical = fs::canonical(requested, error);
    if (error || canonical.empty()) {
        throw std::invalid_argument("cannot canonicalize --spool-dir");
    }
    if (canonical == canonical.root_path()) {
        throw std::invalid_argument("--spool-dir must not be a filesystem root");
    }
    return canonical;
}

fs::path CreateContainedAttachmentFile(
    const fs::path& spoolDirectory,
    const std::wstring& entryId,
    const ULONG attachmentNumber,
    const std::wstring& safeName,
    WinHandle& file) {
    std::wstring messageToken = entryId;
    if (messageToken.size() > 32) {
        messageToken.resize(32);
    }
    if (messageToken.empty()) {
        messageToken = L"unknown";
    }

    const std::wstring base =
        L"rag_" + std::to_wstring(GetCurrentProcessId()) + L"_" +
        messageToken + L"_" + std::to_wstring(attachmentNumber) + L"_" + safeName;
    for (unsigned int attempt = 0; attempt < 1000; ++attempt) {
        const std::wstring candidateName = attempt == 0
            ? base
            : std::to_wstring(attempt) + L"_" + base;
        const fs::path candidate = spoolDirectory / candidateName;
        if (candidate.parent_path() != spoolDirectory) {
            throw std::runtime_error("attachment path escaped spool directory");
        }

        HANDLE rawFile = CreateFileW(
            candidate.c_str(),
            GENERIC_WRITE,
            0,
            nullptr,
            CREATE_NEW,
            FILE_ATTRIBUTE_TEMPORARY,
            nullptr);
        if (rawFile != INVALID_HANDLE_VALUE) {
            file.reset(rawFile);
            return candidate;
        }
        const DWORD createError = GetLastError();
        if (createError != ERROR_FILE_EXISTS && createError != ERROR_ALREADY_EXISTS) {
            throw std::runtime_error("CreateFileW failed for attachment spool file");
        }
    }
    throw std::runtime_error("could not allocate a unique attachment spool file");
}

struct AttachmentInfo {
    std::wstring name;
    std::uint64_t size = 0;
    std::wstring contentType;
    std::wstring tempPath;
};

bool SaveAttachmentStream(
    IStream* stream,
    const fs::path& spoolDirectory,
    const std::wstring& entryId,
    const ULONG attachmentNumber,
    const std::wstring& safeName,
    const std::uint64_t byteLimit,
    AttachmentInfo& info,
    Counters& counters) {
    std::vector<BYTE> buffer(kAttachmentReadBufferBytes);
    WinHandle output;
    fs::path target;
    try {
        target = CreateContainedAttachmentFile(
            spoolDirectory,
            entryId,
            attachmentNumber,
            safeName,
            output);
    } catch (const std::exception&) {
        ++counters.errors;
        std::wcerr << L"  [error] could not create contained attachment spool file\n";
        return false;
    }

    std::uint64_t total = 0;
    bool complete = false;
    while (true) {
        const std::uint64_t remaining = byteLimit >= total ? byteLimit - total : 0;
        const ULONG request = static_cast<ULONG>(std::min<std::uint64_t>(
            buffer.size(),
            remaining == 0 ? 1 : remaining + 1));
        ULONG bytesRead = 0;
        const HRESULT readStatus = stream->Read(buffer.data(), request, &bytesRead);
        if (FAILED(readStatus)) {
            ++counters.errors;
            std::wcerr << L"  [error] IStream::Read(attachment) failed: "
                       << HResultText(readStatus) << L'\n';
            break;
        }
        if (bytesRead == 0) {
            complete = true;
            break;
        }
        if (bytesRead > remaining) {
            break;
        }

        DWORD bytesWritten = 0;
        if (!WriteFile(output.get(), buffer.data(), bytesRead, &bytesWritten, nullptr) ||
            bytesWritten != bytesRead) {
            ++counters.errors;
            std::wcerr << L"  [error] WriteFile(attachment spool) failed\n";
            break;
        }
        total += bytesRead;
    }

    output.reset();
    if (!complete) {
        DeleteFileW(target.c_str());
        return false;
    }

    info.size = total;
    info.tempPath = target.wstring();
    counters.attachmentBytesSaved += total;
    ++counters.attachmentsSaved;
    return true;
}

struct BodyPreview {
    std::wstring text;
    bool available = false;
    bool truncated = false;
    HRESULT status = MAPI_E_NOT_FOUND;
};

BodyPreview ReadUnicodeBodyStream(IMessage* message, const std::size_t maximumChars) {
    BodyPreview result;
    LPUNKNOWN unknown = nullptr;
    result.status = message->OpenProperty(
        PR_BODY_W,
        const_cast<LPIID>(&IID_IStream),
        0,
        0,
        &unknown);
    if (FAILED(result.status) || unknown == nullptr) {
        return result;
    }

    ComPtr<IUnknown> unknownGuard(unknown);
    auto* stream = reinterpret_cast<IStream*>(unknown);
    const std::size_t requestedChars = maximumChars + 1;
    std::vector<wchar_t> buffer(requestedChars, L'\0');
    ULONG bytesRead = 0;
    const HRESULT readResult = stream->Read(
        buffer.data(),
        static_cast<ULONG>(requestedChars * sizeof(wchar_t)),
        &bytesRead);
    if (FAILED(readResult)) {
        result.status = readResult;
        return result;
    }

    const std::size_t charsRead = bytesRead / sizeof(wchar_t);
    std::size_t meaningfulChars = charsRead;
    while (meaningfulChars != 0 && buffer[meaningfulChars - 1] == L'\0') {
        --meaningfulChars;
    }
    result.truncated = meaningfulChars > maximumChars;
    const std::size_t charsToKeep = std::min(meaningfulChars, maximumChars);
    result.text.assign(buffer.data(), charsToKeep);
    while (!result.text.empty() && result.text.back() == L'\0') {
        result.text.pop_back();
    }
    result.available = true;
    result.status = S_OK;
    return result;
}

BodyPreview ReadAnsiBodyStream(IMessage* message, const std::size_t maximumChars) {
    BodyPreview result;
    LPUNKNOWN unknown = nullptr;
    result.status = message->OpenProperty(
        PR_BODY_A,
        const_cast<LPIID>(&IID_IStream),
        0,
        0,
        &unknown);
    if (FAILED(result.status) || unknown == nullptr) {
        return result;
    }

    ComPtr<IUnknown> unknownGuard(unknown);
    auto* stream = reinterpret_cast<IStream*>(unknown);
    const std::size_t requestedBytes = maximumChars + 1;
    std::vector<char> buffer(requestedBytes, '\0');
    ULONG bytesRead = 0;
    const HRESULT readResult = stream->Read(
        buffer.data(),
        static_cast<ULONG>(requestedBytes),
        &bytesRead);
    if (FAILED(readResult)) {
        result.status = readResult;
        return result;
    }

    std::size_t meaningfulBytes = bytesRead;
    while (meaningfulBytes != 0 && buffer[meaningfulBytes - 1] == '\0') {
        --meaningfulBytes;
    }
    result.truncated = meaningfulBytes > maximumChars;
    const std::size_t bytesToKeep = std::min<std::size_t>(meaningfulBytes, maximumChars);
    result.text = FromAnsi(buffer.data(), bytesToKeep);
    while (!result.text.empty() && result.text.back() == L'\0') {
        result.text.pop_back();
    }
    result.available = true;
    result.status = S_OK;
    return result;
}

BodyPreview ReadBodyProperty(IMessage* message, const std::size_t maximumChars) {
    SizedSPropTagArray(2, tags) = {2, {PR_BODY_W, PR_BODY_A}};
    ULONG propertyCount = 0;
    LPSPropValue rawProperties = nullptr;
    const HRESULT status = message->GetProps(
        reinterpret_cast<LPSPropTagArray>(&tags),
        0,
        &propertyCount,
        &rawProperties);
    MapiBufferPtr<SPropValue> properties(rawProperties);

    BodyPreview result;
    result.status = status;
    if (FAILED(status) || rawProperties == nullptr) {
        return result;
    }

    const SPropValue* body = FindProperty(rawProperties, propertyCount, PROP_ID(PR_BODY), PT_UNICODE);
    if (body == nullptr) {
        body = FindProperty(rawProperties, propertyCount, PROP_ID(PR_BODY), PT_STRING8);
    }
    result.text = PropertyString(body);
    if (body == nullptr) {
        return result;
    }
    result.available = true;
    result.truncated = result.text.size() > maximumChars;
    if (result.truncated) {
        result.text.resize(maximumChars);
    }
    result.status = S_OK;
    return result;
}

BodyPreview ReadBody(IMessage* message, const std::size_t maximumChars) {
    if (maximumChars == 0) {
        BodyPreview skipped;
        skipped.available = true;
        skipped.text.clear();
        skipped.status = S_OK;
        return skipped;
    }

    BodyPreview body = ReadUnicodeBodyStream(message, maximumChars);
    if (!body.available) {
        body = ReadAnsiBodyStream(message, maximumChars);
    }
    if (!body.available) {
        body = ReadBodyProperty(message, maximumChars);
    }
    return body;
}

void PrintMapiError(const std::wstring& operation, HRESULT status, Counters& counters);

std::vector<AttachmentInfo> ReadAttachments(
    IMessage* message,
    const std::wstring& entryId,
    const Options& options,
    Counters& counters) {
    std::vector<AttachmentInfo> result;
    LPMAPITABLE rawTable = nullptr;
    const HRESULT tableStatus = message->GetAttachmentTable(MAPI_DEFERRED_ERRORS, &rawTable);
    if (tableStatus == MAPI_E_NO_SUPPORT || tableStatus == MAPI_E_NOT_FOUND) {
        return result;
    }
    if (FAILED(tableStatus) || rawTable == nullptr) {
        PrintMapiError(L"IMessage::GetAttachmentTable", tableStatus, counters);
        return result;
    }
    ComPtr<IMAPITable> table(rawTable);

    SizedSPropTagArray(1, columns) = {1, {PR_ATTACH_NUM}};
    const HRESULT columnsStatus = table->SetColumns(
        reinterpret_cast<LPSPropTagArray>(&columns),
        0);
    if (FAILED(columnsStatus)) {
        PrintMapiError(L"IMAPITable::SetColumns(attachments)", columnsStatus, counters);
        return result;
    }

    while (true) {
        LPSRowSet rawRows = nullptr;
        const HRESULT queryStatus = table->QueryRows(
            static_cast<LONG>(kQueryBatchSize),
            0,
            &rawRows);
        RowSetPtr rows(rawRows);
        if (FAILED(queryStatus)) {
            PrintMapiError(L"IMAPITable::QueryRows(attachments)", queryStatus, counters);
            return result;
        }
        if (rawRows == nullptr || rawRows->cRows == 0) {
            return result;
        }

        for (ULONG rowIndex = 0; rowIndex < rawRows->cRows; ++rowIndex) {
            const SRow& row = rawRows->aRow[rowIndex];
            const SPropValue* numberProperty = FindProperty(
                row.lpProps,
                row.cValues,
                PROP_ID(PR_ATTACH_NUM),
                PT_LONG);
            if (numberProperty == nullptr || numberProperty->Value.l < 0) {
                ++counters.errors;
                std::wcerr << L"  [error] attachment row has no valid PR_ATTACH_NUM\n";
                continue;
            }
            const ULONG attachmentNumber = static_cast<ULONG>(numberProperty->Value.l);

            LPATTACH rawAttachment = nullptr;
            const HRESULT openStatus = message->OpenAttach(
                attachmentNumber,
                nullptr,
                0,
                &rawAttachment);
            if (FAILED(openStatus) || rawAttachment == nullptr) {
                PrintMapiError(L"IMessage::OpenAttach", openStatus, counters);
                continue;
            }
            ComPtr<IAttach> attachment(rawAttachment);

            SizedSPropTagArray(8, tags) = {
                8,
                {
                    PR_ATTACH_METHOD,
                    PR_ATTACH_SIZE,
                    PR_ATTACH_LONG_FILENAME_W,
                    PR_ATTACH_LONG_FILENAME_A,
                    PR_ATTACH_FILENAME_W,
                    PR_ATTACH_FILENAME_A,
                    PR_ATTACH_MIME_TAG_W,
                    PR_ATTACH_MIME_TAG_A
                }
            };
            ULONG propertyCount = 0;
            LPSPropValue rawProperties = nullptr;
            const HRESULT propertiesStatus = attachment->GetProps(
                reinterpret_cast<LPSPropTagArray>(&tags),
                0,
                &propertyCount,
                &rawProperties);
            MapiBufferPtr<SPropValue> properties(rawProperties);
            if (FAILED(propertiesStatus) || rawProperties == nullptr) {
                PrintMapiError(L"IAttach::GetProps", propertiesStatus, counters);
                continue;
            }

            AttachmentInfo info;
            info.name = PropertyStringById(
                rawProperties,
                propertyCount,
                PROP_ID(PR_ATTACH_LONG_FILENAME));
            if (info.name.empty()) {
                info.name = PropertyStringById(
                    rawProperties,
                    propertyCount,
                    PROP_ID(PR_ATTACH_FILENAME));
            }
            if (info.name.empty()) {
                info.name = L"attachment-" + std::to_wstring(attachmentNumber) + L".bin";
            }
            const std::wstring safeSpoolName = SanitizeAttachmentName(info.name);
            info.contentType = PropertyStringById(
                rawProperties,
                propertyCount,
                PROP_ID(PR_ATTACH_MIME_TAG));

            const SPropValue* sizeProperty = FindProperty(
                rawProperties,
                propertyCount,
                PROP_ID(PR_ATTACH_SIZE),
                PT_LONG);
            if (sizeProperty != nullptr && sizeProperty->Value.l > 0) {
                info.size = static_cast<std::uint64_t>(sizeProperty->Value.l);
            }
            const SPropValue* methodProperty = FindProperty(
                rawProperties,
                propertyCount,
                PROP_ID(PR_ATTACH_METHOD),
                PT_LONG);
            const LONG method = methodProperty == nullptr ? 0 : methodProperty->Value.l;

            ++counters.attachments;
            bool saved = false;
            if (method == ATTACH_BY_VALUE && !options.spoolDirectory.empty() &&
                options.maxAttachmentBytes != 0) {
                LPUNKNOWN rawStream = nullptr;
                const HRESULT streamStatus = attachment->OpenProperty(
                    PR_ATTACH_DATA_BIN,
                    const_cast<LPIID>(&IID_IStream),
                    0,
                    0,
                    &rawStream);
                if (SUCCEEDED(streamStatus) && rawStream != nullptr) {
                    ComPtr<IUnknown> streamObject(rawStream);
                    auto* stream = reinterpret_cast<IStream*>(rawStream);

                    STATSTG streamStat = {};
                    if (SUCCEEDED(stream->Stat(&streamStat, STATFLAG_NONAME)) &&
                        streamStat.cbSize.HighPart >= 0) {
                        info.size =
                            (static_cast<std::uint64_t>(streamStat.cbSize.HighPart) << 32) |
                            streamStat.cbSize.LowPart;
                    }

                    const std::uint64_t totalRemaining = options.maxTotalAttachmentBytes == 0
                        ? std::numeric_limits<std::uint64_t>::max()
                        : (counters.attachmentBytesSaved >= options.maxTotalAttachmentBytes
                            ? 0
                            : options.maxTotalAttachmentBytes - counters.attachmentBytesSaved);
                    const std::uint64_t allowed = std::min(
                        options.maxAttachmentBytes,
                        totalRemaining);
                    if (allowed != 0 && info.size <= allowed) {
                        saved = SaveAttachmentStream(
                            stream,
                            options.spoolDirectory,
                            entryId,
                            attachmentNumber,
                            safeSpoolName,
                            allowed,
                            info,
                            counters);
                    }
                } else if (streamStatus != MAPI_E_NOT_FOUND && streamStatus != MAPI_E_NO_SUPPORT) {
                    PrintMapiError(L"IAttach::OpenProperty(PR_ATTACH_DATA_BIN)", streamStatus, counters);
                }
            }
            if (!saved) {
                ++counters.attachmentsSkipped;
            }
            result.push_back(std::move(info));
        }
    }
}

void PrintMapiError(const std::wstring& operation, const HRESULT status, Counters& counters) {
    ++counters.errors;
    std::wcerr << L"  [error] " << operation << L" failed: " << HResultText(status) << L'\n';
}

void PrintMessage(
    IMsgStore* store,
    const SBinary& tableEntryId,
    const std::wstring& inheritedStoreId,
    const std::wstring& storeName,
    const std::wstring& folderEntryId,
    const std::wstring& folderPath,
    const Options& options,
    Counters& counters) {
    ULONG objectType = 0;
    LPUNKNOWN rawObject = nullptr;
    const HRESULT openStatus = store->OpenEntry(
        tableEntryId.cb,
        reinterpret_cast<LPENTRYID>(tableEntryId.lpb),
        nullptr,
        0,
        &objectType,
        &rawObject);
    if (FAILED(openStatus) || rawObject == nullptr || objectType != MAPI_MESSAGE) {
        if (rawObject != nullptr) {
            rawObject->Release();
        }
        PrintMapiError(L"IMsgStore::OpenEntry(message)", openStatus, counters);
        return;
    }

    ComPtr<IUnknown> object(rawObject);
    auto* message = reinterpret_cast<IMessage*>(rawObject);

    SizedSPropTagArray(28, tags) = {
        28,
        {
            PR_ENTRYID,
            PR_STORE_ENTRYID,
            PR_SUBJECT_W,
            PR_SUBJECT_A,
            PR_SENDER_NAME_W,
            PR_SENDER_NAME_A,
            PR_SENDER_SMTP_ADDRESS_W,
            PR_SENDER_SMTP_ADDRESS_A,
            PR_SENDER_EMAIL_ADDRESS_W,
            PR_SENDER_EMAIL_ADDRESS_A,
            PR_SENT_REPRESENTING_NAME_W,
            PR_SENT_REPRESENTING_NAME_A,
            PR_SENT_REPRESENTING_SMTP_ADDRESS_W,
            PR_SENT_REPRESENTING_SMTP_ADDRESS_A,
            PR_SENT_REPRESENTING_EMAIL_ADDRESS_W,
            PR_SENT_REPRESENTING_EMAIL_ADDRESS_A,
            PR_DISPLAY_TO_W,
            PR_DISPLAY_TO_A,
            PR_DISPLAY_CC_W,
            PR_DISPLAY_CC_A,
            PR_CLIENT_SUBMIT_TIME,
            PR_MESSAGE_DELIVERY_TIME,
            PR_LAST_MODIFICATION_TIME,
            PR_INTERNET_MESSAGE_ID_W,
            PR_INTERNET_MESSAGE_ID_A,
            PR_CONVERSATION_ID,
            PR_CONVERSATION_INDEX,
            PR_SEARCH_KEY
        }
    };
    ULONG propertyCount = 0;
    LPSPropValue rawProperties = nullptr;
    const HRESULT propertiesStatus = message->GetProps(
        reinterpret_cast<LPSPropTagArray>(&tags),
        0,
        &propertyCount,
        &rawProperties);
    MapiBufferPtr<SPropValue> properties(rawProperties);
    if (FAILED(propertiesStatus) || rawProperties == nullptr) {
        PrintMapiError(L"IMessage::GetProps", propertiesStatus, counters);
        return;
    }

    const SPropValue* entryProperty = FindProperty(
        rawProperties,
        propertyCount,
        PROP_ID(PR_ENTRYID),
        PT_BINARY);
    const SPropValue* storeProperty = FindProperty(
        rawProperties,
        propertyCount,
        PROP_ID(PR_STORE_ENTRYID),
        PT_BINARY);
    const std::wstring entryId = entryProperty == nullptr
        ? BinaryToHex(&tableEntryId)
        : BinaryToHex(PropertyBinary(entryProperty));
    const std::wstring storeId = storeProperty == nullptr
        ? inheritedStoreId
        : BinaryToHex(PropertyBinary(storeProperty));
    const std::wstring subject = PropertyStringById(
        rawProperties,
        propertyCount,
        PROP_ID(PR_SUBJECT));
    std::wstring senderName = PropertyStringById(
        rawProperties,
        propertyCount,
        PROP_ID(PR_SENDER_NAME));
    if (senderName.empty()) {
        senderName = PropertyStringById(
            rawProperties,
            propertyCount,
            PROP_ID(PR_SENT_REPRESENTING_NAME));
    }
    std::wstring senderEmail = PropertyStringById(
        rawProperties,
        propertyCount,
        PROP_ID(PR_SENDER_SMTP_ADDRESS));
    if (senderEmail.empty()) {
        senderEmail = PropertyStringById(
            rawProperties,
            propertyCount,
            PROP_ID(PR_SENT_REPRESENTING_SMTP_ADDRESS));
    }
    if (senderEmail.empty()) {
        senderEmail = PropertyStringById(
            rawProperties,
            propertyCount,
            PROP_ID(PR_SENDER_EMAIL_ADDRESS));
    }
    if (senderEmail.empty()) {
        senderEmail = PropertyStringById(
            rawProperties,
            propertyCount,
            PROP_ID(PR_SENT_REPRESENTING_EMAIL_ADDRESS));
    }
    const std::wstring displayTo = PropertyStringById(
        rawProperties,
        propertyCount,
        PROP_ID(PR_DISPLAY_TO));
    const std::wstring displayCc = PropertyStringById(
        rawProperties,
        propertyCount,
        PROP_ID(PR_DISPLAY_CC));
    const std::optional<std::wstring> sentAt = PropertyUtcIso8601(
        rawProperties,
        propertyCount,
        PROP_ID(PR_CLIENT_SUBMIT_TIME));
    const std::optional<std::wstring> receivedAt = PropertyUtcIso8601(
        rawProperties,
        propertyCount,
        PROP_ID(PR_MESSAGE_DELIVERY_TIME));
    const std::optional<std::wstring> modifiedAt = PropertyUtcIso8601(
        rawProperties,
        propertyCount,
        PROP_ID(PR_LAST_MODIFICATION_TIME));
    const std::wstring internetMessageId = PropertyStringById(
        rawProperties,
        propertyCount,
        PROP_ID(PR_INTERNET_MESSAGE_ID));
    std::wstring conversationId = BinaryPropertyHex(
        rawProperties,
        propertyCount,
        PROP_ID(PR_CONVERSATION_ID));
    if (conversationId.empty()) {
        conversationId = BinaryPropertyHex(
            rawProperties,
            propertyCount,
            PROP_ID(PR_CONVERSATION_INDEX));
    }
    if (conversationId.empty()) {
        conversationId = BinaryPropertyHex(
            rawProperties,
            propertyCount,
            PROP_ID(PR_SEARCH_KEY));
    }
    const BodyPreview body = ReadBody(message, options.bodyPreviewChars);
    const std::vector<AttachmentInfo> attachments = ReadAttachments(
        message,
        entryId,
        options,
        counters);

    ++counters.messages;
    if (options.jsonl) {
        std::wcout << L"{\"store_id\":\"" << JsonEscape(storeId)
                   << L"\",\"store_name\":\"" << JsonEscape(storeName)
                   << L"\",\"entry_id\":\"" << JsonEscape(entryId)
                   << L"\",\"folder_entry_id\":\"" << JsonEscape(folderEntryId)
                   << L"\",\"folder_path\":\"" << JsonEscape(folderPath)
                   << L"\",\"subject\":\"" << JsonEscape(subject)
                   << L"\",\"body\":\"" << JsonEscape(body.available ? body.text : L"")
                   << L"\",\"body_available\":" << (body.available ? L"true" : L"false")
                   << L",\"body_truncated\":" << (body.truncated ? L"true" : L"false")
                   << L",\"sender_name\":\"" << JsonEscape(senderName)
                   << L"\",\"sender_email\":\"" << JsonEscape(senderEmail)
                   << L"\",\"to\":\"" << JsonEscape(displayTo)
                   << L"\",\"cc\":\"" << JsonEscape(displayCc)
                   << L"\",\"sent_at\":";
        PrintJsonNullableString(sentAt);
        std::wcout << L",\"received_at\":";
        PrintJsonNullableString(receivedAt);
        std::wcout << L",\"modified_at\":";
        PrintJsonNullableString(modifiedAt);
        std::wcout << L",\"internet_message_id\":\"" << JsonEscape(internetMessageId)
                   << L"\",\"conversation_id\":\"" << JsonEscape(conversationId)
                   << L"\",\"attachments\":[";
        for (std::size_t index = 0; index < attachments.size(); ++index) {
            if (index != 0) {
                std::wcout << L',';
            }
            const AttachmentInfo& attachment = attachments[index];
            std::wcout << L"{\"name\":\"" << JsonEscape(attachment.name)
                       << L"\",\"size\":" << attachment.size
                       << L",\"content_type\":\"" << JsonEscape(attachment.contentType)
                       << L"\",\"temp_path\":\"" << JsonEscape(attachment.tempPath)
                       << L"\"}";
        }
        std::wcout << L"]}\n";
        return;
    }

    const std::wstring printableSubject = CompactForConsole(subject);
    const std::wstring printableBody = CompactForConsole(body.text);
    std::wcout << L"    [message " << counters.messages << L"]\n"
               << L"      Folder:  " << folderPath << L'\n'
               << L"      Subject: " << (printableSubject.empty() ? L"<empty>" : printableSubject) << L'\n'
               << L"      EntryID: " << entryId << L'\n'
               << L"      StoreID: " << storeId << L'\n'
               << L"      Attachments: " << attachments.size() << L'\n';
    if (body.available) {
        std::wcout << L"      Body:    " << (printableBody.empty() ? L"<empty>" : printableBody)
                   << (body.truncated ? L" ... <truncated>" : L"") << L'\n';
    } else {
        std::wcout << L"      Body:    <unavailable, " << HResultText(body.status) << L">\n";
    }
}

void EnumerateMessages(
    IMsgStore* store,
    IMAPIFolder* folder,
    const std::wstring& storeId,
    const std::wstring& storeName,
    const std::wstring& folderEntryId,
    const std::wstring& folderPath,
    const Options& options,
    Counters& counters) {
    if (!IsWithinLimit(counters.messages, options.maxMessages)) {
        return;
    }

    LPMAPITABLE rawTable = nullptr;
    const HRESULT tableStatus = folder->GetContentsTable(MAPI_DEFERRED_ERRORS, &rawTable);
    if (tableStatus == MAPI_E_NO_SUPPORT || tableStatus == MAPI_E_NOT_FOUND) {
        return;
    }
    if (FAILED(tableStatus) || rawTable == nullptr) {
        PrintMapiError(L"IMAPIFolder::GetContentsTable", tableStatus, counters);
        return;
    }
    ComPtr<IMAPITable> table(rawTable);

    SizedSPropTagArray(3, columns) = {
        3,
        {PR_ENTRYID, PR_MESSAGE_CLASS_W, PR_MESSAGE_CLASS_A}
    };
    const HRESULT columnsStatus = table->SetColumns(
        reinterpret_cast<LPSPropTagArray>(&columns),
        0);
    if (FAILED(columnsStatus)) {
        PrintMapiError(L"IMAPITable::SetColumns(contents)", columnsStatus, counters);
        return;
    }

    while (IsWithinLimit(counters.messages, options.maxMessages)) {
        std::size_t request = kQueryBatchSize;
        if (options.maxMessages != 0) {
            request = std::min(request, options.maxMessages - counters.messages);
        }

        LPSRowSet rawRows = nullptr;
        const HRESULT queryStatus = table->QueryRows(static_cast<LONG>(request), 0, &rawRows);
        RowSetPtr rows(rawRows);
        if (FAILED(queryStatus)) {
            PrintMapiError(L"IMAPITable::QueryRows(contents)", queryStatus, counters);
            return;
        }
        if (rawRows == nullptr || rawRows->cRows == 0) {
            return;
        }

        for (ULONG rowIndex = 0;
             rowIndex < rawRows->cRows && IsWithinLimit(counters.messages, options.maxMessages);
             ++rowIndex) {
            const SRow& row = rawRows->aRow[rowIndex];
            const std::wstring messageClass = PropertyStringById(
                row.lpProps,
                row.cValues,
                PROP_ID(PR_MESSAGE_CLASS));
            if (!IsEmailMessageClass(messageClass)) {
                ++counters.skippedNonMail;
                continue;
            }
            const SPropValue* entryProperty = FindProperty(
                row.lpProps,
                row.cValues,
                PROP_ID(PR_ENTRYID),
                PT_BINARY);
            const SBinary* entryId = PropertyBinary(entryProperty);
            if (entryId == nullptr) {
                ++counters.errors;
                std::wcerr << L"  [error] contents row has no PR_ENTRYID\n";
                continue;
            }
            PrintMessage(
                store,
                *entryId,
                storeId,
                storeName,
                folderEntryId,
                folderPath,
                options,
                counters);
        }
    }
}

struct ChildFolder {
    std::vector<BYTE> entryId;
    std::wstring displayName;
};

std::vector<ChildFolder> ReadChildFolders(IMAPIFolder* folder, Counters& counters) {
    std::vector<ChildFolder> children;
    LPMAPITABLE rawTable = nullptr;
    const HRESULT tableStatus = folder->GetHierarchyTable(MAPI_DEFERRED_ERRORS, &rawTable);
    if (tableStatus == MAPI_E_NO_SUPPORT || tableStatus == MAPI_E_NOT_FOUND) {
        return children;
    }
    if (FAILED(tableStatus) || rawTable == nullptr) {
        PrintMapiError(L"IMAPIFolder::GetHierarchyTable", tableStatus, counters);
        return children;
    }
    ComPtr<IMAPITable> table(rawTable);

    SizedSPropTagArray(2, columns) = {2, {PR_ENTRYID, PR_DISPLAY_NAME_W}};
    const HRESULT columnsStatus = table->SetColumns(
        reinterpret_cast<LPSPropTagArray>(&columns),
        0);
    if (FAILED(columnsStatus)) {
        PrintMapiError(L"IMAPITable::SetColumns(hierarchy)", columnsStatus, counters);
        return children;
    }

    while (true) {
        LPSRowSet rawRows = nullptr;
        const HRESULT queryStatus = table->QueryRows(static_cast<LONG>(kQueryBatchSize), 0, &rawRows);
        RowSetPtr rows(rawRows);
        if (FAILED(queryStatus)) {
            PrintMapiError(L"IMAPITable::QueryRows(hierarchy)", queryStatus, counters);
            return children;
        }
        if (rawRows == nullptr || rawRows->cRows == 0) {
            return children;
        }

        for (ULONG rowIndex = 0; rowIndex < rawRows->cRows; ++rowIndex) {
            const SRow& row = rawRows->aRow[rowIndex];
            const SPropValue* entryProperty = FindProperty(
                row.lpProps,
                row.cValues,
                PROP_ID(PR_ENTRYID),
                PT_BINARY);
            const SBinary* entryId = PropertyBinary(entryProperty);
            if (entryId == nullptr) {
                ++counters.errors;
                std::wcerr << L"  [error] hierarchy row has no PR_ENTRYID\n";
                continue;
            }
            const SPropValue* nameProperty = FindProperty(
                row.lpProps,
                row.cValues,
                PROP_ID(PR_DISPLAY_NAME),
                PT_UNICODE);
            children.push_back({
                CopyBinary(entryId),
                CompactForConsole(PropertyString(nameProperty))
            });
        }
    }
}

std::wstring ReadFolderEntryId(IMAPIFolder* folder) {
    SizedSPropTagArray(1, tags) = {1, {PR_ENTRYID}};
    ULONG propertyCount = 0;
    LPSPropValue rawProperties = nullptr;
    const HRESULT status = folder->GetProps(
        reinterpret_cast<LPSPropTagArray>(&tags),
        0,
        &propertyCount,
        &rawProperties);
    MapiBufferPtr<SPropValue> properties(rawProperties);
    if (FAILED(status) || rawProperties == nullptr) {
        return L"<missing>";
    }
    return BinaryToHex(PropertyBinary(FindProperty(
        rawProperties,
        propertyCount,
        PROP_ID(PR_ENTRYID),
        PT_BINARY)));
}

void EnumerateFolder(
    IMsgStore* store,
    IMAPIFolder* folder,
    const std::wstring& storeId,
    const std::wstring& storeName,
    const std::wstring& folderPath,
    const Options& options,
    Counters& counters,
    std::unordered_set<std::wstring>& visitedFolderIds) {
    if (!IsWithinLimit(counters.folders, options.maxFolders) ||
        !IsWithinLimit(counters.messages, options.maxMessages)) {
        return;
    }

    const std::wstring folderId = ReadFolderEntryId(folder);
    if (folderId.empty() || folderId == L"<missing>") {
        ++counters.errors;
        std::wcerr << L"  [error] folder has no valid PR_ENTRYID\n";
        return;
    }
    const std::wstring visitKey = storeId + L":" + folderId;
    if (!visitedFolderIds.insert(visitKey).second) {
        return;
    }

    ++counters.folders;
    std::wostream& diagnostic = options.jsonl ? std::wcerr : std::wcout;
    diagnostic << L"  [folder " << counters.folders << L"] " << folderPath << L'\n'
               << L"    EntryID: " << folderId << L'\n';
    EnumerateMessages(
        store,
        folder,
        storeId,
        storeName,
        folderId,
        folderPath,
        options,
        counters);

    if (!IsWithinLimit(counters.folders, options.maxFolders) ||
        !IsWithinLimit(counters.messages, options.maxMessages)) {
        return;
    }

    const std::vector<ChildFolder> children = ReadChildFolders(folder, counters);
    for (const ChildFolder& child : children) {
        if (!IsWithinLimit(counters.folders, options.maxFolders) ||
            !IsWithinLimit(counters.messages, options.maxMessages)) {
            break;
        }
        if (child.entryId.empty()) {
            continue;
        }

        ULONG objectType = 0;
        LPUNKNOWN rawObject = nullptr;
        const HRESULT openStatus = store->OpenEntry(
            static_cast<ULONG>(child.entryId.size()),
            reinterpret_cast<LPENTRYID>(const_cast<BYTE*>(child.entryId.data())),
            nullptr,
            0,
            &objectType,
            &rawObject);
        if (FAILED(openStatus) || rawObject == nullptr || objectType != MAPI_FOLDER) {
            if (rawObject != nullptr) {
                rawObject->Release();
            }
            PrintMapiError(L"IMsgStore::OpenEntry(folder)", openStatus, counters);
            continue;
        }

        ComPtr<IUnknown> object(rawObject);
        auto* childFolder = reinterpret_cast<IMAPIFolder*>(rawObject);
        const std::wstring childName = child.displayName.empty() ? L"<unnamed>" : child.displayName;
        const std::wstring childPath = folderPath + L"/" + childName;
        EnumerateFolder(
            store,
            childFolder,
            storeId,
            storeName,
            childPath,
            options,
            counters,
            visitedFolderIds);
    }
}

void EnumerateStore(
    IMAPISession* session,
    const SBinary& storeEntryId,
    const std::wstring& displayName,
    const bool isDefault,
    const Options& options,
    Counters& counters) {
    LPMDB rawStore = nullptr;
    const HRESULT openStatus = session->OpenMsgStore(
        0,
        storeEntryId.cb,
        reinterpret_cast<LPENTRYID>(storeEntryId.lpb),
        nullptr,
        MDB_NO_DIALOG,
        &rawStore);
    if (FAILED(openStatus) || rawStore == nullptr) {
        PrintMapiError(L"IMAPISession::OpenMsgStore", openStatus, counters);
        return;
    }
    ComPtr<IMsgStore> store(rawStore);

    const std::wstring storeId = BinaryToHex(&storeEntryId);
    ++counters.stores;
    std::wostream& diagnostic = options.jsonl ? std::wcerr : std::wcout;
    diagnostic << L"[store " << counters.stores << L"] "
               << (displayName.empty() ? L"<unnamed>" : displayName)
               << (isDefault ? L" [default]" : L"") << L'\n'
               << L"  StoreID: " << storeId << L'\n';

    // Start at the IPM subtree so the default probe traverses user-visible mailbox
    // content instead of internal Finder/Common Views nodes under the store root.
    const std::vector<BYTE> ipmSubtreeEntryId = ReadBinaryProperty(
        store.get(),
        PR_IPM_SUBTREE_ENTRYID);

    ULONG objectType = 0;
    LPUNKNOWN rawRoot = nullptr;
    const HRESULT rootStatus = store->OpenEntry(
        static_cast<ULONG>(ipmSubtreeEntryId.size()),
        ipmSubtreeEntryId.empty()
            ? nullptr
            : reinterpret_cast<LPENTRYID>(const_cast<BYTE*>(ipmSubtreeEntryId.data())),
        nullptr,
        0,
        &objectType,
        &rawRoot);
    if (FAILED(rootStatus) || rawRoot == nullptr || objectType != MAPI_FOLDER) {
        if (rawRoot != nullptr) {
            rawRoot->Release();
        }
        PrintMapiError(L"IMsgStore::OpenEntry(root)", rootStatus, counters);
        return;
    }

    ComPtr<IUnknown> root(rawRoot);
    auto* rootFolder = reinterpret_cast<IMAPIFolder*>(rawRoot);
    std::unordered_set<std::wstring> visitedFolderIds;
    const std::wstring rootName = displayName.empty() ? L"<store-root>" : displayName;
    EnumerateFolder(
        store.get(),
        rootFolder,
        storeId,
        displayName,
        rootName,
        options,
        counters,
        visitedFolderIds);
}

void EnumerateStores(IMAPISession* session, const Options& options, Counters& counters) {
    LPMAPITABLE rawTable = nullptr;
    const HRESULT tableStatus = session->GetMsgStoresTable(0, &rawTable);
    if (FAILED(tableStatus) || rawTable == nullptr) {
        throw std::runtime_error("IMAPISession::GetMsgStoresTable failed");
    }
    ComPtr<IMAPITable> table(rawTable);

    SizedSPropTagArray(3, columns) = {
        3,
        {PR_ENTRYID, PR_DISPLAY_NAME_W, PR_DEFAULT_STORE}
    };
    const HRESULT columnsStatus = table->SetColumns(
        reinterpret_cast<LPSPropTagArray>(&columns),
        0);
    if (FAILED(columnsStatus)) {
        throw std::runtime_error("IMAPITable::SetColumns(stores) failed");
    }

    while (IsWithinLimit(counters.stores, options.maxStores)) {
        LPSRowSet rawRows = nullptr;
        const HRESULT queryStatus = table->QueryRows(static_cast<LONG>(kQueryBatchSize), 0, &rawRows);
        RowSetPtr rows(rawRows);
        if (FAILED(queryStatus)) {
            throw std::runtime_error("IMAPITable::QueryRows(stores) failed");
        }
        if (rawRows == nullptr || rawRows->cRows == 0) {
            return;
        }

        for (ULONG rowIndex = 0;
             rowIndex < rawRows->cRows && IsWithinLimit(counters.stores, options.maxStores);
             ++rowIndex) {
            const SRow& row = rawRows->aRow[rowIndex];
            const SPropValue* entryProperty = FindProperty(
                row.lpProps,
                row.cValues,
                PROP_ID(PR_ENTRYID),
                PT_BINARY);
            const SBinary* entryId = PropertyBinary(entryProperty);
            if (entryId == nullptr) {
                ++counters.errors;
                std::wcerr << L"[error] store row has no PR_ENTRYID\n";
                continue;
            }
            const SPropValue* nameProperty = FindProperty(
                row.lpProps,
                row.cValues,
                PROP_ID(PR_DISPLAY_NAME),
                PT_UNICODE);
            const SPropValue* defaultProperty = FindProperty(
                row.lpProps,
                row.cValues,
                PROP_ID(PR_DEFAULT_STORE),
                PT_BOOLEAN);
            const bool isDefault = defaultProperty != nullptr && defaultProperty->Value.b != FALSE;
            const std::wstring displayName = PropertyString(nameProperty);
            if (!ContainsCaseInsensitive(displayName, options.storeContains)) {
                continue;
            }
            EnumerateStore(
                session,
                *entryId,
                displayName,
                isDefault,
                options,
                counters);
            if (!IsWithinLimit(counters.messages, options.maxMessages)) {
                return;
            }
        }
    }
}

std::size_t ParseLimit(
    const wchar_t* value,
    const wchar_t* optionName,
    const std::size_t maximum = kMaximumCliLimit) {
    if (value == nullptr || *value == L'\0' || *value == L'-') {
        throw std::invalid_argument("invalid numeric option");
    }
    std::size_t consumed = 0;
    const unsigned long long parsed = std::stoull(value, &consumed, 10);
    if (value[consumed] != L'\0' || parsed > maximum) {
        std::wcerr << optionName << L" must be between 0 and " << maximum << L'\n';
        throw std::invalid_argument("numeric option out of range");
    }
    return static_cast<std::size_t>(parsed);
}

std::uint64_t ParseByteLimit(const wchar_t* value, const wchar_t* optionName) {
    if (value == nullptr || *value == L'\0' || *value == L'-') {
        throw std::invalid_argument("invalid byte limit");
    }
    std::size_t consumed = 0;
    const unsigned long long parsed = std::stoull(value, &consumed, 10);
    if (value[consumed] != L'\0' || parsed > kMaximumCliByteLimit) {
        std::wcerr << optionName << L" must be between 0 and "
                   << kMaximumCliByteLimit << L'\n';
        throw std::invalid_argument("byte limit out of range");
    }
    return static_cast<std::uint64_t>(parsed);
}

void PrintUsage(std::wostream& output) {
    output
        << L"NativeMapiProbe (x64, read-only Extended MAPI)\n\n"
        << L"Usage: NativeMapiProbe.exe [options]\n\n"
        << L"  --max-stores N          Stores to open; 0 = unlimited (default: 0)\n"
        << L"  --max-folders N         Folders to open globally; 0 = unlimited (default: 100)\n"
        << L"  --max-messages N        Messages to open globally; 0 = unlimited (default: 20)\n"
        << L"  --store-contains TEXT   Only open stores whose display name contains TEXT\n"
        << L"  --body-preview-chars N  Body characters to read/print; 0 disables preview\n"
        << L"                          (default: 240; maximum: 4000000)\n"
        << L"  --spool-dir PATH        Explicit directory for bounded by-value attachment\n"
        << L"                          extraction; omitted = metadata only/no file writes\n"
        << L"  --max-attachment-bytes N\n"
        << L"                          Per-attachment hard cap; 0 disables extraction\n"
        << L"                          (default: 67108864 / 64 MiB)\n"
        << L"  --max-total-attachment-bytes N\n"
        << L"                          Per-process extracted-byte cap; 0 = unlimited\n"
        << L"                          (default: 0; per-attachment cap still applies)\n"
        << L"  --jsonl                 UTF-8 JSON object per message on stdout;\n"
        << L"                          diagnostics and summary go to stderr\n"
        << L"  --help                  Show this text\n";
}

bool ParseOptions(const int argc, wchar_t* argv[], Options& options) {
    for (int index = 1; index < argc; ++index) {
        const std::wstring argument = argv[index];
        if (argument == L"--help" || argument == L"-h" || argument == L"/?") {
            PrintUsage(std::wcout);
            return false;
        }
        if (argument == L"--jsonl") {
            options.jsonl = true;
            continue;
        }

        if (argument == L"--store-contains" || argument == L"--spool-dir") {
            if (index + 1 >= argc) {
                std::wcerr << L"Missing value for " << argument << L'\n';
                throw std::invalid_argument("missing option value");
            }
            if (argument == L"--store-contains") {
                options.storeContains = argv[++index];
            } else {
                options.spoolDirectory = argv[++index];
            }
            continue;
        }

        if (index + 1 >= argc) {
            std::wcerr << L"Missing value for " << argument << L'\n';
            throw std::invalid_argument("missing option value");
        }
        const wchar_t* value = argv[++index];
        if (argument == L"--max-stores") {
            options.maxStores = ParseLimit(value, argument.c_str());
        } else if (argument == L"--max-folders") {
            options.maxFolders = ParseLimit(value, argument.c_str());
        } else if (argument == L"--max-messages") {
            options.maxMessages = ParseLimit(value, argument.c_str());
        } else if (argument == L"--body-preview-chars") {
            options.bodyPreviewChars = ParseLimit(
                value,
                argument.c_str(),
                kMaximumBodyChars);
        } else if (argument == L"--max-attachment-bytes") {
            options.maxAttachmentBytes = ParseByteLimit(value, argument.c_str());
        } else if (argument == L"--max-total-attachment-bytes") {
            options.maxTotalAttachmentBytes = ParseByteLimit(value, argument.c_str());
        } else {
            std::wcerr << L"Unknown option: " << argument << L'\n';
            throw std::invalid_argument("unknown option");
        }
    }
    options.spoolDirectory = PrepareSpoolDirectory(options.spoolDirectory);
    return true;
}

}  // namespace

int wmain(const int argc, wchar_t* argv[]) {
    // Keep redirected output deterministic too: the default "C" locale makes
    // std::wcout enter fail state at the first Cyrillic subject/body character.
    _setmode(_fileno(stdout), _O_U8TEXT);
    _setmode(_fileno(stderr), _O_U8TEXT);
    SetConsoleOutputCP(CP_UTF8);
    SetConsoleCP(CP_UTF8);

    Options options;
    try {
        if (!ParseOptions(argc, argv, options)) {
            return 0;
        }
    } catch (const std::exception&) {
        PrintUsage(std::wcerr);
        return 64;
    }

    try {
        MapiRuntime mapi;
        LPMAPISESSION rawSession = nullptr;
        const FLAGS logonFlags =
            MAPI_EXTENDED |
            MAPI_USE_DEFAULT |
            MAPI_NEW_SESSION |
            MAPI_NO_MAIL;
        const HRESULT logonStatus = MAPILogonEx(
            0,
            nullptr,
            nullptr,
            logonFlags,
            &rawSession);
        if (FAILED(logonStatus) || rawSession == nullptr) {
            std::wcerr << L"MAPILogonEx(default profile) failed: "
                       << HResultText(logonStatus) << L'\n';
            return 3;
        }
        ComPtr<IMAPISession> session(rawSession);

        std::wostream& diagnostic = options.jsonl ? std::wcerr : std::wcout;
        diagnostic << L"Logged on to the default Outlook profile via Extended MAPI.\n"
                   << L"Mode: x64, no UI, no send/write calls.\n"
                   << L"Limits: stores=" << options.maxStores
                   << L", folders=" << options.maxFolders
                   << L", messages=" << options.maxMessages
                   << L", body-preview=" << options.bodyPreviewChars << L" chars"
                   << L", attachment=" << options.maxAttachmentBytes << L" bytes"
                   << L", attachment-total=" << options.maxTotalAttachmentBytes << L" bytes"
                   << (options.spoolDirectory.empty()
                       ? L", attachment-mode=metadata-only"
                       : L", attachment-mode=bounded-spool")
                   << (options.storeContains.empty()
                       ? L""
                       : L", store-filter=\"" + options.storeContains + L"\"")
                   << L"\n\n";

        Counters counters;
        EnumerateStores(session.get(), options, counters);

        diagnostic << L"\nSummary: stores=" << counters.stores
                   << L", folders=" << counters.folders
                   << L", messages=" << counters.messages
                   << L", skipped_non_mail=" << counters.skippedNonMail
                   << L", attachments=" << counters.attachments
                   << L", attachments-saved=" << counters.attachmentsSaved
                   << L", attachments-skipped=" << counters.attachmentsSkipped
                   << L", attachment-bytes-saved=" << counters.attachmentBytesSaved
                   << L", recoverable-errors=" << counters.errors << L'\n';

        session->Logoff(0, 0, 0);
        return counters.errors == 0 ? 0 : 1;
    } catch (const std::exception& error) {
        std::wcerr << L"Fatal: " << FromAnsi(error.what()) << L'\n';
        return 4;
    }
}
