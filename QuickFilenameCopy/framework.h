#pragma once


#include "targetver.h"
#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#define STRICT 1
#define STRICT_TYPED_ITEMIDS
#include <windows.h>
#include <tlhelp32.h>
#include <psapi.h>
#include <shellapi.h> 
#include <shlwapi.h> 
#include <combaseapi.h> 
#include <shlobj.h>

#pragma warning(disable : 4471)
//#import <mshtml.tlb> no_implementation
//#import <shdocvw.dll> no_implementation

#include <exdisp.h>

#include <unordered_set> 
#include <string>
#include <tuple> 
#if __cpp_lib_format
#include <format>
#else
#define FMT_HEADER_ONLY 
#include "fmt/format.h"

namespace std {
    using namespace ::fmt;
}
#endif
