// SqrlTray.cpp : Implementation of WinMain


#include "stdafx.h"
#include "resource.h"
#include "SqrlTray_i.h"
#include "xdlldata.h"
#include "SqrlSettingsDlg.h"
#include "SqrlAboutDlg.h"

//Include novaPDF headers
#include "..\novaSDK\include\novaOptions.h"

//NovaPdfOptions
#include "..\novaSDK\include\novapi.h"

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
HINSTANCE							g_hInstance = NULL;
HICON								g_hGhIconMain = NULL;
UINT								g_uShellRestart;
HWND								g_hWndTray = NULL;

BOOL	DoTrayIconAndTip(DWORD);
BOOL	DoBalloonTip(LPTSTR sBalloonTitle, LPTSTR sBalloonTip, UINT uBalloonTimeout);
void	ContextMenu(HWND);
bool	IsVista();

HRESULT	InitNovaOptions();
HRESULT	SwitchToSqrlProfile();
void	UnInitNovaOptions();

// Callbacks
BOOL CALLBACK TrayProc(HWND, UINT, WPARAM, LPARAM);

// Data 
const int WM_EX_MESSAGE = (WM_APP + 1);
const int SQRL_TRAY_ICON_ID = 1966;

//register File Saved message for novaPDF Printer
const UINT  wm_Nova_StartDoc = RegisterWindowMessageW(MSG_NOVAPDF2_STARTDOC);
const UINT  wm_Nova_EndDoc = RegisterWindowMessageW(MSG_NOVAPDF2_ENDDOC);
const UINT  wm_Nova_FileSaved = RegisterWindowMessageW(MSG_NOVAPDF2_FILESAVED);
const UINT  wm_Nova_PrintError = RegisterWindowMessageW(MSG_NOVAPDF2_PRINTERROR);

#define PRINTER_NAME        L"novaPDF SDK 8"
#define SQRL_PROFILE		L"Sqrl Profile"
#define PROFILE_IS_PUBLIC   0	// private profile

LPWSTR  m_wsProfileSqrl = NULL;
LPWSTR  m_wsDefaultProfile = NULL;
LPWSTR  m_strProfileId = NULL;

INovaPdfOptions80 *m_novaOptions = NULL;

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

	if (SUCCEEDED(InitNovaOptions())){

		if (SUCCEEDED(SwitchToSqrlProfile())){

			DoTrayIconAndTip(NIM_ADD);

			DoBalloonTip(_T("SqrlTray ready...!"), _T(" "), 10);

			DoTrayIconAndTip(NIM_MODIFY);

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
			}
		}
	}

	//cleanup: 
	UnInitNovaOptions();

	if (m_novaOptions)
		m_novaOptions->Release();
	m_novaOptions = NULL;

	DoTrayIconAndTip(NIM_DELETE);

#ifdef _NEWBTIP
	if (g_hAppMainICO)
		DestroyIcon(g_hAppMainICO);
#endif
	DestroyWindow(g_hWndTray);

	if (g_hGhIconMain)
		DestroyIcon(g_hGhIconMain);

	CoUninitialize();

	return 0;
}

HRESULT InitNovaOptions(){

	HRESULT hr = CoCreateInstance(__uuidof(NovaPdfOptions80), NULL, CLSCTX_INPROC_SERVER, __uuidof(INovaPdfOptions80), (LPVOID*)&m_novaOptions);
	if (!SUCCEEDED(hr)){
		DoBalloonTip(_T("SqrlTray"), _T("Error getting novaPDF options - check installation"), 10);
		return hr;
	}

	// initialize the NovaPdfOptions object to use with a printer licensed for SDK
	// if you have an application license for novaPDF SDK, 
	// pass the license key to the Initialize() function
	hr = m_novaOptions->Initialize(PRINTER_NAME, L"");	//TODO - license key...
	if (!SUCCEEDED(hr)){
		DoBalloonTip(_T("SqrlTray"), _T("Error initializing novaPDF printer - check installation"), 10);
		return hr;
	}

	hr = m_novaOptions->RegisterEventWindow((LONG)g_hWndTray);
	if (!SUCCEEDED(hr)){
		DoBalloonTip(_T("SqrlTray"), _T("Error registering for novaPDF events - check installation"), 10);
		return hr;
	}

	return S_OK;
}

void UnInitNovaOptions(){
	if (m_novaOptions) {

		//unregister events
		m_novaOptions->UnRegisterEventWindow();

		//restore default profile
		if (m_wsDefaultProfile)
			m_novaOptions->SetActiveProfile(m_wsDefaultProfile);

		//delete newly created profile
		if (m_wsProfileSqrl)
			m_novaOptions->DeleteProfile(m_wsProfileSqrl);

		//free memory
		if (m_wsProfileSqrl){
			CoTaskMemFree(m_wsProfileSqrl);
			m_wsProfileSqrl = NULL;
		}
		if (m_wsDefaultProfile){
			CoTaskMemFree(m_wsDefaultProfile);
			m_wsDefaultProfile = NULL;
		}

		//restore default printer
		m_novaOptions->RestoreDefaultPrinter();
	}
}

HRESULT SwitchToSqrlProfile(){

	//set novaPDF default printer 
	ATLASSERT(m_novaOptions);

	m_novaOptions->SetDefaultPrinter();

	HRESULT hr;
	hr = m_novaOptions->GetActiveProfile(&m_wsDefaultProfile);	// may be NULL

	//create a new profile with default settings
	hr = m_novaOptions->AddProfile(SQRL_PROFILE, PROFILE_IS_PUBLIC, &m_wsProfileSqrl);
	switch(hr){
		case S_OK:
			ATLASSERT(m_wsProfileSqrl);
			hr = m_novaOptions->LoadProfile(m_wsProfileSqrl);
			ATLASSERT(SUCCEEDED(hr) && m_wsProfileSqrl);
			break;
		case NV_NOT_INITIALIZED:// Initialize was not called
		case NV_SERVICE_ERROR:// cannot connect to novaPDF Server service
		case NV_PROFILE_ERROR://cannot read default profile
		case NV_PROFILE_SAVE_ERROR://cannot save new profile
			// maybe it already exists... just load it ...
			//TODO
			DoBalloonTip(_T("SqrlTray"), _T("Error creating Sqrl profile..."), 10);
			return S_FALSE;
	}

	// get the directory of the application to save the PDF fles there
	WCHAR szExeDirectory[MAX_PATH] = L"";
	WCHAR drive[MAX_PATH], dir[MAX_PATH];
	GetModuleFileNameW(NULL, szExeDirectory, MAX_PATH);
	_wsplitpath(szExeDirectory, drive, dir, NULL, NULL);
	wcscpy(szExeDirectory, drive);
	wcscat(szExeDirectory, dir);

	// disable the "Save PDF file as" prompt
	m_novaOptions->SetOptionLong(NOVAPDF_SAVE_PROMPT_TYPE, PROMPT_SAVE_NONE);

	// set generated Pdf files destination folder ("c:\")
	m_novaOptions->SetOptionLong(NOVAPDF_SAVE_LOCATION, LOCATION_TYPE_LOCAL);
	m_novaOptions->SetOptionLong(NOVAPDF_SAVE_FOLDER_TYPE, SAVEFOLDER_CUSTOM);
	m_novaOptions->SetOptionString(NOVAPDF_SAVE_FOLDER, szExeDirectory);

	// set output file name
	m_novaOptions->SetOptionString(NOVAPDF_SAVE_FILE_NAME, L"Sqrl.pdf");

	// if file exists in the destination folder, append a counter to the end of the file name
	m_novaOptions->SetOptionLong(NOVAPDF_SAVE_FILEEXIST_ACTION, FILE_CONFLICT_STRATEGY_AUTONUMBER_NEW);

	m_novaOptions->SetOptionBool(NOVAPDF_ACTION_DEFAULT_VIEWER, FALSE);

	//save profile changes
	m_novaOptions->SaveProfile();

	// switch active profile to Sqrl...
	m_novaOptions->SetActiveProfile(m_wsProfileSqrl);

	return S_OK;
}


BOOL CALLBACK TrayProc(HWND hDlg, UINT uiMsg, WPARAM wParam, LPARAM lParam)
{
	if (uiMsg == wm_Nova_FileSaved){
		ATLTRACE("MSG_NOVAPDF2_FILESAVED....");

		// REST call here....

		// bring up browser here.....

		//TODO delete generated PDF ....
	}
	else if (uiMsg == wm_Nova_StartDoc){
		ATLTRACE("MSG_NOVAPDF2_STARTDOC....");
	}
	else if (uiMsg == wm_Nova_EndDoc){
		ATLTRACE("MSG_NOVAPDF2_ENDDOC....");
	}
	else if (uiMsg == wm_Nova_PrintError){
		ATLTRACE("MSG_NOVAPDF2_PRINTERROR....");
		switch (wParam)
		{
		case ERROR_MSG_TEMP_FILE:
			DoBalloonTip(_T("SqrlTray"), _T("Error saving temporary file on printer server"), 10);
			break;
		case ERROR_MSG_LIC_INFO:
			DoBalloonTip(_T("SqrlTray"), _T("Error reading license information"), 10);
			break;
		case ERROR_MSG_SAVE_PDF:
			DoBalloonTip(_T("SqrlTray"), _T("Error saving PDF file"), 10);
			break;
		case ERROR_MSG_JOB_CANCELED:
			DoBalloonTip(_T("SqrlTray"), _T("Print job was canceled"), 10);
			break;
		}

	}
	else{

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

bool IsVista()	
{
	/*TODO OSVERSIONINFO osVer;
	memset(&osVer, 0, sizeof(OSVERSIONINFO));
	osVer.dwOSVersionInfoSize = sizeof(OSVERSIONINFO);
	GetVersionEx(&osVer);
	return (osVer.dwMajorVersion >= 6);*/
	return false;
}