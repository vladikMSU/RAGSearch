cmake_minimum_required(VERSION 3.20)

project(NativeMapiProbe LANGUAGES CXX)

if(NOT WIN32)
    message(FATAL_ERROR "NativeMapiProbe is Windows-only (Extended MAPI).")
endif()

if(NOT CMAKE_SIZEOF_VOID_P EQUAL 8)
    message(FATAL_ERROR "NativeMapiProbe must be configured for x64 to match 64-bit Outlook.")
endif()

add_executable(NativeMapiProbe main.cpp)

set(MAPI_INCLUDE_DIR "" CACHE PATH "Directory containing MAPIX.h/MAPITags.h")
if(NOT MAPI_INCLUDE_DIR)
    find_path(MAPI_INCLUDE_DIR
        NAMES MAPIX.h mapix.h
        HINTS
            "$ENV{MAPI_INCLUDE_DIR}"
            "C:/Program Files/Microsoft SDKs/Office/14.0/Include")
endif()
if(NOT MAPI_INCLUDE_DIR)
    message(FATAL_ERROR
        "Extended MAPI headers were not found. Install MAPIStubLibrary headers "
        "or configure -DMAPI_INCLUDE_DIR=<path-to-MAPIStubLibrary/include>.")
endif()

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
