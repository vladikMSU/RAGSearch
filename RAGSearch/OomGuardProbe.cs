using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace RAGSearch
{
    /// <summary>
    /// A deliberately unsafe, one-message Outlook Object Model negative control.
    /// Production indexing never calls this class.  It exists only so a human can
    /// compare a documented protected OOM getter with the Extended MAPI path.
    /// </summary>
    internal sealed class OomGuardProbe
    {
        private const int EFail = unchecked((int)0x80004005);
        private const int EAbort = unchecked((int)0x80004004);
        private const int EAccessDenied = unchecked((int)0x80070005);
        private const int MapiENotSupported = unchecked((int)0x80040102);

        private readonly Outlook.Application application;
        private readonly string logPath;

        public OomGuardProbe(Outlook.Application application)
        {
            this.application = application ?? throw new ArgumentNullException("application");
            logPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "RAGSearch",
                "oom-guard-probe.log");
        }

        /// <summary>
        /// Must be called synchronously by the Outlook/WinForms click handler.  In
        /// particular, do not put an await or a Timer between locating the message
        /// and reading SenderEmailAddress: the protected getter itself is the probe.
        /// </summary>
        public string Run(Action<string> reportPhase)
        {
            Outlook.MailItem mail = null;
            var runId = Guid.NewGuid().ToString("N");
            Log(string.Format(
                "BEGIN run={0}; tid={1}; apartment={2}",
                runId,
                Thread.CurrentThread.ManagedThreadId,
                Thread.CurrentThread.GetApartmentState()));

            try
            {
                Report(
                    reportPhase,
                    "DEBUG OOM [1/3]: ищу одно реальное MailItem без чтения адреса или текста...");

                string source;
                mail = FindOneMail(out source);
                if (mail == null)
                {
                    const string notFound =
                        "DEBUG OOM result: NO_MAIL — в selection/current folder/Inbox/stores не найден MailItem.";
                    Log("END run=" + runId + "; result=NO_MAIL");
                    return notFound;
                }

                Log("run=" + runId + "; mail located; source=" + source);
                Report(
                    reportPhase,
                    "DEBUG OOM [2/3]: MailItem найден (" + source +
                    "). Сейчас читаю protected MailItem.SenderEmailAddress; ожидайте Guard...");

                try
                {
                    Log("run=" + runId +
                        "; PROTECTED_GETTER_BEGIN member=MailItem.SenderEmailAddress");

                    // Documented protected address-information property.  The
                    // returned address is deliberately discarded and never logged.
                    // https://learn.microsoft.com/office/vba/outlook/how-to/security/
                    // protected-properties-and-methods
                    var protectedAddress = mail.SenderEmailAddress;
                    var returnedNonEmpty = !string.IsNullOrEmpty(protectedAddress);
                    protectedAddress = null;

                    Log(string.Format(
                        "run={0}; PROTECTED_GETTER_RETURN member=MailItem.SenderEmailAddress; nonempty={1}",
                        runId,
                        returnedNonEmpty));
                    Log("END run=" + runId + "; result=GETTER_RETURNED");
                    return string.Format(
                        "DEBUG OOM [3/3] result=GETTER_RETURNED, HRESULT=none, value_nonempty={0}. " +
                        "Это Allow, уже действующий временный grant либо отсутствие prompt; сам OOM их не различает. Адрес не сохранён.",
                        returnedNonEmpty);
                }
                catch (COMException ex)
                {
                    var hresult = ex.ErrorCode;
                    // Outlook documents MAPI_E_NOT_SUPPORTED for a denied
                    // protected OOM call. Keep E_ABORT/E_ACCESSDENIED because
                    // older/current builds and providers can surface either.
                    var result = hresult == MapiENotSupported ||
                                 hresult == EFail ||
                                 hresult == EAbort ||
                                 hresult == EAccessDenied
                        ? "DENY_OR_BLOCKED"
                        : "COM_EXCEPTION";
                    Log(string.Format(
                        "run={0}; PROTECTED_GETTER_THROW member=MailItem.SenderEmailAddress; result={1}; hresult=0x{2:X8}",
                        runId,
                        result,
                        hresult));
                    Log("END run=" + runId + "; result=" + result);
                    return string.Format(
                        "DEBUG OOM [3/3] result={0}, HRESULT=0x{1:X8}. " +
                        "Значение адреса не было прочитано и не логировалось.",
                        result,
                        hresult);
                }
            }
            catch (COMException ex)
            {
                Log(string.Format(
                    "END run={0}; result=LOCATE_COM_EXCEPTION; hresult=0x{1:X8}",
                    runId,
                    ex.ErrorCode));
                return string.Format(
                    "DEBUG OOM result=LOCATE_COM_EXCEPTION, HRESULT=0x{0:X8}; protected getter ещё не вызывался.",
                    ex.ErrorCode);
            }
            catch (Exception ex)
            {
                Log("END run=" + runId + "; result=ERROR; type=" + ex.GetType().FullName);
                return "DEBUG OOM result=ERROR, type=" + ex.GetType().Name +
                    "; protected getter мог не выполниться. Смотрите oom-guard-probe.log.";
            }
            finally
            {
                ComRelease.Final(mail);
            }
        }

        private Outlook.MailItem FindOneMail(out string source)
        {
            Outlook.Explorer explorer = null;
            Outlook.Selection selection = null;
            object candidate = null;
            try
            {
                explorer = application.ActiveExplorer();
                if (explorer == null)
                {
                    source = null;
                    return null;
                }
                selection = explorer.Selection;
                if (selection == null || selection.Count < 1)
                {
                    source = null;
                    return null;
                }

                candidate = selection[1];
                var selectedMail = candidate as Outlook.MailItem;
                if (selectedMail != null)
                {
                    candidate = null;
                    source = "current selection";
                    return selectedMail;
                }

                source = null;
                return null;
            }
            finally
            {
                ComRelease.Final(candidate);
                ComRelease.Final(selection);
                ComRelease.Final(explorer);
            }
        }

        private void Report(Action<string> reportPhase, string message)
        {
            Log("PHASE " + message);
            if (reportPhase != null)
            {
                reportPhase(message);
            }
        }

        private void Log(string message)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(logPath));
                File.AppendAllText(
                    logPath,
                    DateTime.UtcNow.ToString("o") + " " + message + Environment.NewLine);
            }
            catch (Exception)
            {
                // Diagnostics must not affect the probe result.
            }
        }
    }
}
