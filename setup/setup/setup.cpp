// Tablacus Script Control 64 Installer (C)2014 Gaku
// MIT Lisence
// Visual C++ 2008 Express Edition SP1
// Windows SDK v7.0
// http://www.eonet.ne.jp/~gakana/tablacus/

#include "stdafx.h"
#include "setup.h"

#define MAX_LOADSTRING 100

// Global Variables:
HINSTANCE hInst;								// current instance
TCHAR szTitle[MAX_LOADSTRING];					// The title bar text
TCHAR szWindowClass[MAX_LOADSTRING];			// the main window class name

// Forward declarations of functions included in this code module:
ATOM				MyRegisterClass(HINSTANCE hInstance);
BOOL				InitInstance(HINSTANCE, int);
LRESULT CALLBACK	WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK	About(HWND, UINT, WPARAM, LPARAM);

int APIENTRY _tWinMain(HINSTANCE hInstance,
                     HINSTANCE hPrevInstance,
                     LPTSTR    lpCmdLine,
                     int       nCmdShow)
{
	UNREFERENCED_PARAMETER(hPrevInstance);
	UNREFERENCED_PARAMETER(lpCmdLine);

 	// TODO: Place code here.
	MSG msg;

	// Initialize global strings
	LoadString(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
	LoadString(hInstance, IDC_SETUP, szWindowClass, MAX_LOADSTRING);
	MyRegisterClass(hInstance);

	// Perform application initialization:
	if (!InitInstance (hInstance, nCmdShow))
	{
		return FALSE;
	}

	// Main message loop:
	while (GetMessage(&msg, NULL, 0, 0))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}

	return (int) msg.wParam;
}



//
//  FUNCTION: MyRegisterClass()
//
//  PURPOSE: Registers the window class.
//
//  COMMENTS:
//
//    This function and its usage are only necessary if you want this code
//    to be compatible with Win32 systems prior to the 'RegisterClassEx'
//    function that was added to Windows 95. It is important to call this function
//    so that the application will get 'well formed' small icons associated
//    with it.
//
ATOM MyRegisterClass(HINSTANCE hInstance)
{
	WNDCLASSEX wcex;

	wcex.cbSize = sizeof(WNDCLASSEX);

	wcex.style			= CS_HREDRAW | CS_VREDRAW;
	wcex.lpfnWndProc	= WndProc;
	wcex.cbClsExtra		= 0;
	wcex.cbWndExtra		= 0;
	wcex.hInstance		= hInstance;
	wcex.hIcon			= LoadIcon(hInstance, MAKEINTRESOURCE(IDI_SETUP));
	wcex.hCursor		= LoadCursor(NULL, IDC_ARROW);
	wcex.hbrBackground	= (HBRUSH)(COLOR_WINDOW+1);
	wcex.lpszMenuName	= MAKEINTRESOURCE(IDC_SETUP);
	wcex.lpszClassName	= szWindowClass;
	wcex.hIconSm		= LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_SETUP));

	return RegisterClassEx(&wcex);
}

//
//   FUNCTION: InitInstance(HINSTANCE, int)
//
//   PURPOSE: Saves instance handle and creates main window
//
//   COMMENTS:
//
//        In this function, we save the instance handle in a global variable and
//        create and display the main program window.
//
BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
   HWND hWnd;

   hInst = hInstance; // Store instance handle in our global variable

   hWnd = CreateWindow(szWindowClass, szTitle, WS_OVERLAPPEDWINDOW,
      CW_USEDEFAULT, 0, 400, 200, NULL, NULL, hInstance, NULL);

   if (!hWnd)
   {
      return FALSE;
   }
   ShowWindow(hWnd, nCmdShow);
   UpdateWindow(hWnd);

   return TRUE;
}

//
//  FUNCTION: WndProc(HWND, UINT, WPARAM, LPARAM)
//
//  PURPOSE:  Processes messages for the main window.
//
//  WM_COMMAND	- process the application menu
//  WM_PAINT	- Paint the main window
//  WM_DESTROY	- post a quit message and return
//
//
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	int wmId, wmEvent;
	PAINTSTRUCT ps;
	HDC hdc;
	WCHAR szSys32[MAX_PATH * 2];
	WCHAR szRegSvr[MAX_PATH * 2];
	WCHAR szPath[MAX_PATH * 2];
	WCHAR szPath2[MAX_PATH * 2];
	HWND hInstall, hUninstall;
	SHELLEXECUTEINFO sei = { sizeof(SHELLEXECUTEINFO) };

	switch (message)
	{
	case WM_CREATE:
		hInstall = CreateWindow(
			_T("BUTTON"),
			_T("Install"),
			WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
			40, 20,
			300, 40,
			hWnd,
			(HMENU)(INT_PTR)IDM_INSTALL,
			hInst,
			NULL
		);
		hUninstall = CreateWindow(
			_T("BUTTON"),
			_T("Uninstall"),
			WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
			40, 80,
			300, 40,
			hWnd,
			(HMENU)(INT_PTR)IDM_UNINSTALL,
			hInst,
			NULL
		);
		return 0;
	case WM_COMMAND:
		wmId    = LOWORD(wParam);
		wmEvent = HIWORD(wParam);

		GetSystemDirectory(szRegSvr, MAX_PATH * 2);
		PathAppend(szRegSvr, L"regsvr32.exe");
		GetSystemDirectory(szSys32, MAX_PATH * 2);
		PathAppend(szSys32, DLL_FILENAME);
		GetModuleFileName(NULL, szPath, MAX_PATH * 2);
		PathRemoveFileSpec(szPath);
		PathAppend(szPath, DLL_FILENAME);
		sei.lpFile = szRegSvr;
		sei.nShow = SW_SHOWNORMAL;

		// Parse the menu selections:
		switch (wmId)
		{
		case IDM_INSTALL:
			CopyFile(szPath, szSys32, FALSE);
			sei.lpParameters = szSys32;
			ShellExecuteEx(&sei);
			break;
		case IDM_UNINSTALL:
			wsprintf(szPath2, L"/u %s", szSys32);
			sei.lpParameters = szPath2;
			sei.fMask = SEE_MASK_NOCLOSEPROCESS;
			ShellExecuteEx(&sei);
		    ::WaitForSingleObject(sei.hProcess, INFINITE);
			DeleteFile(szSys32);
			break;
		case IDM_EXIT:
			DestroyWindow(hWnd);
			break;
		default:
			return DefWindowProc(hWnd, message, wParam, lParam);
		}
		break;
	case WM_PAINT:
		hdc = BeginPaint(hWnd, &ps);
		// TODO: Add any drawing code here...
		EndPaint(hWnd, &ps);
		break;
	case WM_DESTROY:
		PostQuitMessage(0);
		break;
	default:
		return DefWindowProc(hWnd, message, wParam, lParam);
	}
	return 0;
}

