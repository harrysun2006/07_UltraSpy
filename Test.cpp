#include "StdAfx.h"
#include "MainDlg.h"
#include "Test.h"
#include "FlashSink3.h"
#include "Flash9d.tlh"

#include <oleacc.h>
#include <afxctl.h>

#pragma comment (lib, "oleacc")

#include <mshtml.h>
#include <map>
#include <string>

using namespace std;
using namespace ShockwaveFlashObjects;

void test1()
{
	IWebBrowser2 *m_pInternetExplorer; 
	CoCreateInstance(CLSID_InternetExplorer,NULL,CLSCTX_SERVER,IID_IWebBrowser2,(LPVOID *)&m_pInternetExplorer);
	m_pInternetExplorer->put_Visible(VARIANT_TRUE);

	BSTR Url,Target,PostData,Head;
	VARIANT BstrUrl,BstrTarget,IFlag,BstrPostData,BstrHead;
	V_VT(&BstrUrl)=VT_BSTR;
	V_BSTR(&BstrUrl)=Url=SysAllocString(L"http://www.google.com");
	V_VT(&BstrTarget)=VT_BSTR;
	V_BSTR(&BstrTarget)=Target=SysAllocString(L"_self");
	V_VT(&BstrPostData)=VT_BSTR;
	V_BSTR(&BstrPostData)=PostData=SysAllocString(L"");
	V_VT(&BstrHead)=VT_BSTR;
	V_BSTR(&BstrHead)=Head=SysAllocString(L"");
	V_VT(&IFlag)=VT_I4;
	V_I4(&IFlag)=navNoHistory;

	m_pInternetExplorer->Navigate2(&BstrUrl,&IFlag,&BstrTarget,&BstrPostData,&BstrHead);
	SysFreeString(Url);
	SysFreeString(Target);
	SysFreeString(PostData);
	SysFreeString(Head);
	//MonitorWebBrowser(m_pInternetExplorer);
	m_pInternetExplorer->Release();
}

void test2()
{
	IID id = DIID_DWebBrowserEvents2;
	LPOLESTR strGUID = 0;
	CComBSTR bstrId = NULL;
	// 注意此处StringFromIID的正确用法, 否则会在Debug状态下出错: user breakpoint called from code 0x7c901230
	StringFromIID(id, &strGUID);
	bstrId = strGUID;
	CoTaskMemFree(strGUID);
	USES_CONVERSION;
	g_pMainWin->Log(OLE2CT(bstrId));
}

void test3()
{
	CString strDebug;
	typedef map<string, string>maps;
	typedef pair<string, string>pr;
	maps temp;
	temp.insert(pr("aa","aaaaa"));
	temp.insert(pr("bb","bbbbbb"));
	temp.insert(pr("cc","cccc"));
	maps   tent;
	maps::iterator   it;
	for(it = temp.begin(); it != temp.end(); ++it) tent.insert(*it);
	for(it = tent.begin(); it != tent.end(); ++it)
	{
		strDebug.Format("%s ==> %s\n", (*it).first, (*it).second);
		g_pMainWin->Log(strDebug);
	}
}

void test4()
{
	CString strDebug;
	CFlashSink3* sink = new CFlashSink3();
	IDispatch* pDisp = NULL;
	HRESULT hr = sink->QueryInterface(DIID__IShockwaveFlashEvents, reinterpret_cast<void**>(&pDisp));
	if (SUCCEEDED(hr) && pDisp) g_pMainWin->Log(_T("CFlashSink3::_IShockwaveFlashEvents found.\n"));
	else g_pMainWin->Log(_T("CFlashSink3::_IShockwaveFlashEvents NOT found!"));
}

BOOL CALLBACK EnumIEWindow(HWND hWnd, LPARAM lParam)
{
	TCHAR szClassName[256];
	GetClassName(hWnd,  szClassName,  sizeof(szClassName));

	if (_tcscmp(szClassName, _T("Internet Explorer_Server")) == 0)
	{
		*(HWND*)lParam = hWnd;
		return FALSE;
	}
	return TRUE;
}

void test5()
{
	HWND hWnd = NULL;
	//hWnd = FindWindowEx(NULL, NULL, "IEFrame", NULL);
	EnumChildWindows(GetDesktopWindow(), EnumIEWindow, (LPARAM)&hWnd);
	if (hWnd) 
	{
		// UINT nMsg = RegisterWindowMessage(_T("WM_HTML_GETOBJECT"));
		UINT nMsg = RegisterWindowMessage(_T("WM_HTML_GETOBJECT"));
		LRESULT lRes;
		SendMessageTimeout(hWnd, nMsg, 0L, 0L, SMTO_ABORTIFHUNG, 1000, (DWORD*)&lRes);

		HRESULT hr = S_OK;

		CComPtr<IHTMLDocument2> spDocument = NULL;
		hr = ObjectFromLresult (lRes, IID_IHTMLDocument2, 0 , (LPVOID*)&spDocument);
		if (FAILED(hr) || !spDocument) g_pMainWin->Log(_T("Can NOT found IHTMLDocument2 object!\n"));
		else g_pMainWin->Log(_T("Found IHTMLDocument2 object!\n"));

		CComPtr<IWebBrowser2> spWebBrowser = NULL;
		hr = ObjectFromLresult (lRes, IID_IWebBrowser2, 0 , (LPVOID*)&spWebBrowser);
		if (FAILED(hr) || !spWebBrowser) g_pMainWin->Log(_T("Can NOT found IWebBrowser2 object!\n"));
		else g_pMainWin->Log(_T("Found IWebBrowser2 object!\n"));

		CComPtr<::IServiceProvider> spServiceProvider = NULL;
		hr = ObjectFromLresult (lRes, ::IID_IServiceProvider, 0 , (LPVOID*)&spServiceProvider);
		if (FAILED(hr) || !spServiceProvider) g_pMainWin->Log(_T("Can NOT found IServiceProvider object!\n"));
		else g_pMainWin->Log(_T("Found IServiceProvider object!\n"));

		CComPtr<IWebBrowserApp> spWebBrowserApp = NULL;
		hr = ObjectFromLresult (lRes, IID_IWebBrowserApp, 0 , (LPVOID*)&spWebBrowserApp);
		if (FAILED(hr) || !spWebBrowserApp) g_pMainWin->Log(_T("Can NOT found IWebBrowserApp object!\n"));
		else g_pMainWin->Log(_T("Found IWebBrowserApp object!\n"));
	}
	else
	{
		g_pMainWin->Log(_T("Can NOT found IE window!\n"));
	}
}