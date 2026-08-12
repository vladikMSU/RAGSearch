cmake_minimum_required(VERSION 3.20)

project(NativeMapiProbe LANGUAGES CXX)

if(NOT WIN32)
    message(FATAL_ERROR "NativeMapiProbe is Windows-only (Extended MAPI).")
endif()

if(NOT CMAKE_SIZEOF_VOID_P EQUAL 8)
    message(FATAL_ERROR "NativeMapiProbe must be configured for x64 to match 64-bit Outlook.")
endif()

add_executable(NativeMapiProbe main.cpp)

set_target_properties(NativeMapiProbe PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY "${CMAKE_CURRENT_SOURCE_DIR}/build-direct"
    RUNTIME_OUTPUT_DIRECTORY_DEBUG "${CMAKE_CURRENT_SOURCE_DIR}/build-direct"
    RUNTIME_OUTPUT_DIRECTORY_RELEASE "${CMAKE_CURRENT_SOURCE_DIR}/build-direct"
    RUNTIME_OUTPUT_DIRECTORY_RELWITHDEBINFO "${CMAKE_CURRENT_SOURCE_DIR}/build-direct"
    RUNTIME_OUTPUT_DIRECTORY_MINSIZEREL "${CMAKE_CURRENT_SOURCE_DIR}/build-direct")

set(_MAPI_INCLUDE_DEFAULT
    "${CMAKE_CURRENT_SOURCE_DIR}/../third_party/MAPIStubLibrary/include")
if(DEFINED ENV{MAPI_INCLUDE_DIR} AND NOT "$ENV{MAPI_INCLUDE_DIR}" STREQUAL "")
    set(_MAPI_INCLUDE_DEFAULT "$ENV{MAPI_INCLUDE_DIR}")
endif()
set(MAPI_INCLUDE_DIR
    "${_MAPI_INCLUDE_DEFAULT}"
    CACHE PATH
    "Directory containing MAPIX.h/MAPITags.h")
foreach(_MAPI_HEADER IN ITEMS
        MAPICode.h
        MAPIDefS.h
        MAPIGuid.h
        MAPITags.h
        MAPIUtil.h
        MAPIX.h)
    if(NOT EXISTS "${MAPI_INCLUDE_DIR}/${_MAPI_HEADER}")
        message(FATAL_ERROR
            "Extended MAPI header ${_MAPI_HEADER} was not found. Restore "
            "third_party/MAPIStubLibrary/include or configure "
            "-DMAPI_INCLUDE_DIR=<path-to-MAPIStubLibrary/include>.")
    endif()
endforeach()

target_include_directories(NativeMapiProbe PRIVATE "${MAPI_INCLUDE_DIR}")

target_compile_features(NativeMapiProbe PRIVATE cxx_std_17)
target_compile_definitions(NativeMapiProbe PRIVATE
    UNICODE
    _UNICODE
    WIN32_LEAN_AND_MEAN
    NOMINMAX)

if(MSVC)
    target_compile_options(NativeMapiProbe PRIVATE /W4 /permissive- /utf-8 /EHsc)
endif()

target_link_libraries(NativeMapiProbe PRIVATE MAPI32 Ole32)

add_custom_command(TARGET NativeMapiProbe POST_BUILD
    COMMAND "${CMAKE_COMMAND}" -E copy_if_different
        "${CMAKE_CURRENT_SOURCE_DIR}/../THIRD_PARTY_NOTICES.md"
        "$<TARGET_FILE_DIR:NativeMapiProbe>/THIRD_PARTY_NOTICES.md"
    COMMAND "${CMAKE_COMMAND}" -E copy_if_different
        "${CMAKE_CURRENT_SOURCE_DIR}/../third_party/MAPIStubLibrary/LICENSE"
        "$<TARGET_FILE_DIR:NativeMapiProbe>/MAPIStubLibrary-LICENSE.txt"
    COMMENT "Copying third-party license notices")
