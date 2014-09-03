// IESpy.cpp : Defines the entry point for the application.
//

#include "stdafx.h"
#include "resource.h"
#include "UltraSpy.h"
#include "MainDlg.h"

CComModule _Module;

int APIENTRY WinMain(HINSTANCE hInstance,
                     HINSTANCE hPrevInstance,
                     LPSTR     lpCmdLine,
                     int       nCmdShow)
{
    lpCmdLine = GetCommandLine(); //this line necessary for _ATL_MIN_CRT

    HRESULT hRes = CoInitialize(NULL);
    _ASSERTE(SUCCEEDED(hRes));

    _Module.Init(0, hInstance, &LIBID_ATLLib);

	CMainDlg mainDlg;
	g_pMainWin = &mainDlg;

	mainDlg.Create(NULL);
	mainDlg.ShowWindow(SW_SHOW);

    MSG msg;
    while (GetMessage(&msg, 0, 0, 0))
	{
		if (!IsDialogMessage(mainDlg, &msg))
		{
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}

    _Module.Term();
    CoUninitialize();
    return msg.wParam;
}



