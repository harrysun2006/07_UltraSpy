// MainDlg.cpp: implementation of the CMainDlg class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "MainDlg.h"
#include "Spyer.h"
#include "Test.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

#pragma comment (lib, "ShellHook.lib")
extern BOOL WINAPI InstallShellHookEx(HWND hWnd, UINT uMsg);
extern BOOL WINAPI UnInstallShellHookEx(HWND hWnd);

CMainDlg* g_pMainWin = 0;

using namespace spyer;

class CMyIESink : public CIESink
{
protected:
	virtual void OnDocumentComplete(IDispatch *pDisp, VARIANT *pvarURL);
};

void CMyIESink::OnDocumentComplete(IDispatch *pDisp, VARIANT *pvarURL)
{
	g_pMainWin->Log("MyIESink: Document completed!\n");
}

LRESULT CMainDlg::OnInitDialog(
	UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	CenterWindow();
	bHandled = FALSE;
	return S_OK;
}

LRESULT CMainDlg::OnDestroy(
	UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	PostQuitMessage(0);
	return S_OK;
}

LRESULT CMainDlg::OnCancel(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
{
	DestroyWindow();
	return S_OK;
}

void CMainDlg::SinkIE(HWND hWnd)
{
	CString strDebug;
	CIESink* pSink = new CMyIESink();
	HRESULT hr = CSpyer::Instance()->AddIESink(hWnd, pSink);
	if (SUCCEEDED(hr))
	{
		strDebug.Format("IESink: Add a sink to window[%08X] successfully!\n", hWnd);
	}
	else
	{
		strDebug.Format("IESink: Can NOT add sink to window[%08X]!\n", hWnd);
	}
	Log(strDebug);
}

LRESULT CMainDlg::OnShell(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	CString strDebug, strTitle;
	CWnd* pwnd;
	if(::IsWindow(HWND(lParam)))
	{
		pwnd = CWnd::FromHandle(HWND(lParam));
		pwnd->GetWindowText(strTitle);
	}
	switch (wParam)
	{
	case HSHELL_WINDOWCREATED:
		SinkIE(HWND(lParam));
		strDebug.Format("ShellHook: Window[%08X] - %s created!\n", lParam, strTitle);
		break;
	case HSHELL_WINDOWDESTROYED:
		strDebug.Format("ShellHook: Window[%08X] - %s destroyed!\n", lParam, strTitle);
		break;
	case HSHELL_WINDOWACTIVATED:
		strDebug.Format("ShellHook: Window[%08X] - %s activated!\n", lParam, strTitle);
		break;
	default:
		strDebug.Format("%08X, %08X\n", wParam, lParam);
		break;
	}
	// Log(strDebug);
	return S_OK;
}

LRESULT CMainDlg::OnEnum(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
{
	CMyIESink* pSink = new CMyIESink();
	CSpyer::Instance()->AddIESink(HWND(NULL), pSink);
	return S_OK;
}

LRESULT CMainDlg::OnHook(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
{
	if (InstallShellHookEx(m_hWnd, WM_SHELL))
	{
		Log("Shell hookex installed successfully!\n");
	}
	else
	{
		Log("Can NOT install shell hookex, maybe this window has already installed one!\n");
	}
	return S_OK;
}

LRESULT CMainDlg::OnUnhook(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
{
	UnInstallShellHookEx(m_hWnd);
	return S_OK;
}

LRESULT CMainDlg::OnTest(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
{
	test5();
	return S_OK;
}

LRESULT CMainDlg::OnClear(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
{
	CWindow wndEdit = GetDlgItem(IDC_LOG);
	int len = wndEdit.GetWindowTextLength();
	wndEdit.SendMessage(EM_SETSEL, 0, len);
	wndEdit.SendMessage(EM_REPLACESEL, FALSE, NULL);
	return S_OK;
}

void CMainDlg::Log(LPCTSTR psz)
{
	//CString strDebug;
	//strDebug.Format("%s\n", psz);
	
	CWindow wndEdit = GetDlgItem(IDC_LOG);
	int len = wndEdit.GetWindowTextLength();
	wndEdit.SendMessage(EM_SETSEL, len, len);
	wndEdit.SendMessage(EM_REPLACESEL, FALSE, reinterpret_cast<LPARAM>(psz));
	//wndEdit.SendMessage(EM_REPLACESEL, FALSE, reinterpret_cast<LPARAM>(strDebug.GetBuffer(strDebug.GetLength())));
}
