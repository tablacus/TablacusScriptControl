#pragma once

#include "resource.h"
#include <Shlwapi.h>
#include <shellapi.h>
#pragma comment(lib, "shlwapi.lib")

#ifdef _WIN64
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='amd64' publicKeyToken='6595b64144ccf1df' language='*'\"")
#else
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='x86' publicKeyToken='6595b64144ccf1df' language='*'\"")
#define _2000XP
#endif
//8 or higher
typedef BOOL (WINAPI* LPFNSetDefaultDllDirectories)(__in DWORD DirectoryFlags);
#ifndef LOAD_LIBRARY_SEARCH_SYSTEM32
#define LOAD_LIBRARY_SEARCH_SYSTEM32 0x00000800
#endif
