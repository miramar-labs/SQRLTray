// SqrlTray.cpp : Implementation of WinMain


#include "stdafx.h"
#include "resource.h"
#include "SqrlTray_i.h"
#include "xdlldata.h"
#include "SqrlSettingsDlg.h"
#include "SqrlAboutDlg.h"

using namespace ATL;

class CSqrlTrayModule : public ATL::CAtlExeModuleT< CSqrlTrayModule >
{
public :
	DECLARE_LIBID(LIBID_SqrlTrayLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_SQRLTRAY, "{10BE8F22-4FD0-465B-B9BF-07478B2CCDFA}")
	};

CSqrlTrayModule _AtlModule;

#ifdef _NEWBTIP
HICON								g_hAppMainICO;
#endif

NOTIFYICONDATA						g_nid;
HINSTANCE							g_hInstance;
HICON								g_hGhIconMain;
UINT								g_uShellRestart;
HWND								g_hWndTray;

BOOL	DoTrayIconAndTip(DWORD);
BOOL	DoBalloonTip(LPTSTR sBalloonTitle, LPTSTR sBalloonTip, UINT uBalloonTimeout);
void	ContextMenu(HWND);
bool	IsVista();

// Callbacks
BOOL CALLBACK TrayProc(HWND, UINT, WPARAM, LPARAM);

#define CALL_OK(a) (a == S_OK)

// Data 
const int WM_EX_MESSAGE = (WM_APP + 1);
const int SQRL_TRAY_ICON_ID = 1966;

//
extern "C" int WINAPI _tWinMain(HINSTANCE hInstance, HINSTANCE /*hPrevInstance*/, 
								LPTSTR /*lpCmdLine*/, int nShowCmd)
{
	CoInitialize(NULL);

	if (::FindWindow(NULL, _T("SqrlTray Hidden Wnd"))){
		::MessageBoxA(GetForegroundWindow(), "SqrlTray is already running!", "Sqrl", MB_OK);
		return 0;
	}

	g_hInstance = hInstance;

	g_uShellRestart = RegisterWindowMessage(__TEXT("TaskbarCreated"));

#ifdef _NEWBTIP
	g_hGhIconMain = reinterpret_cast<HICON>(LoadImage(hInstance,
		MAKEINTRESOURCE(IDI_TRAY_APP),
		IMAGE_ICON,
		32,
		32,
		0
		)
		);
	ATLASSERT(g_hAppMainICO);
#endif

	g_hGhIconMain = reinterpret_cast<HICON>(LoadImage(hInstance,
		MAKEINTRESOURCE(IDI_ICONMAIN),
		IMAGE_ICON,
		16,
		16,
		0
		)
		);

	g_hWndTray = CreateDialog(hInstance, MAKEINTRESOURCE(IDD_SQRL_TRAY), NULL, TrayProc);	//hidden window for receiving messages

	ZeroMemory(&g_nid, sizeof(g_nid));

	if (IsVista())
		g_nid.cbSize = sizeof(NOTIFYICONDATA);
	else
		g_nid.cbSize = NOTIFYICONDATA_V2_SIZE;

	g_nid.hWnd = g_hWndTray;
	g_nid.uID = SQRL_TRAY_ICON_ID;
	g_nid.uFlags = NIF_MESSAGE;
	g_nid.uCallbackMessage = WM_EX_MESSAGE;

	DoTrayIconAndTip(NIM_ADD);

	DoBalloonTip(_T("SqrlTray!"), _T(" "), 10);

	DoTrayIconAndTip(NIM_MODIFY);

	ATLTRACE("SqrlTray: ready!");

	MSG msg;
	while (GetMessage(&msg, NULL, 0, 0))
	{
		if (msg.message == WM_QUIT){
			ATLTRACE("WM_QUIT!!");
			break;
		}

		if (!IsDialogMessage(g_hWndTray, &msg))
		{
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
		//ATLTRACE("sqrlTray: processing messages...\n");
	}

	DoTrayIconAndTip(NIM_DELETE);

#ifdef _NEWBTIP
	DestroyIcon(g_hAppMainICO);
#endif
	DestroyWindow(g_hWndTray);
	DestroyIcon(g_hGhIconMain);

	CoUninitialize();

	return 0;
}

BOOL CALLBACK TrayProc(HWND hDlg, UINT uiMsg, WPARAM wParam, LPARAM lParam)
{

	switch (uiMsg)
	{

	case WM_COMMAND:

		switch (LOWORD(wParam))
		{

		case ID_FILE_SETTINGS:
		{
			CSqrlSettingsDlg dlg;
			dlg.DoModal();
		}
		break;

		case ID_FILE_ABOUT:
		{
			CSqrlAboutDlg dlg;
			dlg.DoModal();
		}
		break;

		case IDCANCEL:
		case ID_FILE_EXITSQRLTRAY:
			PostQuitMessage(0);
			return FALSE;
		}
		break;

	case WM_EX_MESSAGE:
		if (wParam == SQRL_TRAY_ICON_ID)
		{
			switch (lParam)
			{

			case WM_RBUTTONUP:
				ContextMenu(hDlg);
				break;

			case NIN_BALLOONUSERCLICK:
			{

			}
			break;
			}
		}
		break;
	}
	return FALSE;
}

BOOL DoTrayIconAndTip(DWORD msg)
{
	g_nid.hIcon = g_hGhIconMain;
	_tcscpy_s(g_nid.szTip, sizeof(g_nid.szTip) / sizeof(TCHAR), _T("SqrlTray"));

	// hIcon is valid..
	g_nid.uFlags |= NIF_ICON;

	// set this as a tooltip..
	g_nid.uFlags |= NIF_TIP;

	// remove any previous balloon tip stuff
	g_nid.uFlags = g_nid.uFlags & ~NIF_INFO;
	g_nid.dwInfoFlags = g_nid.dwInfoFlags &~NIIF_INFO;
	g_nid.dwInfoFlags = g_nid.dwInfoFlags &~NIIF_USER;

	return Shell_NotifyIcon(msg, &g_nid);
}

BOOL DoBalloonTip(LPTSTR sBalloonTitle, LPTSTR sBalloonTip, UINT uBalloonTimeout)
{
	// get rid of any previous flags
	g_nid.uFlags = g_nid.uFlags &~NIF_ICON;
	g_nid.uFlags = g_nid.uFlags &~NIF_TIP;

	// set up baloon tip flags
	g_nid.uFlags |= NIF_INFO;	//use balooon tip

#ifdef _NEWBTIP
	/*
	Vista:   NIIF_LARGE_ICON ==> SM_CXICON x SM_CYICON  (32x32)
	else				 SM_CXSMICON x SM_CYSMICON (16x16)  (reg & green pickles OK)
	*/

	if (IsVista()){
		g_nid.dwInfoFlags = NIIF_USER | NIIF_LARGE_ICON;
		g_nid.hBalloonIcon = g_hAppMainICO;
	}
	else{
		g_nid.dwInfoFlags = NIIF_USER;
		g_nid.uFlags |= NIF_ICON;
		g_nid.hIcon = g_hAppMainICO;
	}
#else
	g_nid.dwInfoFlags = NIIF_INFO;
#endif

	// set timeout and tip text
	ATLASSERT(uBalloonTimeout >= 10 && uBalloonTimeout <= 30);
	g_nid.uTimeout = uBalloonTimeout * 1000;

	_tcsncpy_s(g_nid.szInfo, sizeof(g_nid.szInfo) / sizeof(TCHAR), sBalloonTip, _tcslen(sBalloonTip));

	if (sBalloonTitle)
		_tcsncpy_s(g_nid.szInfoTitle, sizeof(g_nid.szInfoTitle) / sizeof(TCHAR), sBalloonTitle, _tcslen(sBalloonTitle));
	else
		g_nid.szInfoTitle[0] = _T('\0');

	return Shell_NotifyIcon(NIM_MODIFY, &g_nid);
}

void ContextMenu(HWND hwnd)
{

	HMENU hmenu = LoadMenu(g_hInstance, MAKEINTRESOURCE(IDC_SQRL_TRAY));
	HMENU hmnuPopup = GetSubMenu(hmenu, 0);
	SetMenuDefaultItem(hmnuPopup, IDOK, FALSE);

	POINT pt;
	GetCursorPos(&pt);
	SetForegroundWindow(hwnd);
	TrackPopupMenu(hmnuPopup, TPM_LEFTALIGN, pt.x, pt.y, 0, hwnd, NULL);
	SetForegroundWindow(hwnd);

	DestroyMenu(hmnuPopup);
	DestroyMenu(hmenu);
}

bool IsVista()	// logic using this was if !IsVista then XP ... 
{
	/*TODO OSVERSIONINFO osVer;
	memset(&osVer, 0, sizeof(OSVERSIONINFO));
	osVer.dwOSVersionInfoSize = sizeof(OSVERSIONINFO);
	GetVersionEx(&osVer);
	return (osVer.dwMajorVersion >= 6);*/
	return false;
}