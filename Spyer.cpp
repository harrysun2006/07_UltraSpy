// Spyer.cpp : Manage sinks
//

#include "stdafx.h"
#include "Spyer.h"
#include "IESink.h"
#include "FlashSink.h"
#include "IEHelper.h"
#include "Flash9d.tlh"

#include <oleacc.h>
#include <afxctl.h>

#pragma comment (lib, "oleacc")

using namespace ShockwaveFlashObjects;
using namespace iehelper;

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

namespace spyer {

CSpyer* CSpyer::_instance = NULL;

HRESULT CSpyer::AddIESink(IWebBrowser2* pWebBrowser2, CIESink* pSink)
{
	if (!pSink) return E_INVALIDARG;
	// add IE sink to all IE browsers
	if (!pWebBrowser2)
	{
		CComPtr<IShellWindows> spShellWin;
		HRESULT hr = spShellWin.CoCreateInstance(CLSID_ShellWindows);
		if (FAILED(hr)) return hr;

		long nCount = 0;
		// 取得浏览器实例个数(Explorer 和 IExplorer)
		spShellWin->get_Count(&nCount);

		for(int i = 0; i < nCount; i++)
		{
			CComPtr<IDispatch> spDispIE;
			hr = spShellWin->Item(CComVariant((long)i), &spDispIE);
			if (FAILED (hr)) continue;

			CComQIPtr<IWebBrowser2> spWebBrowser2 = spDispIE;
			if (!spWebBrowser2) continue;
			hr = AddIESink(spWebBrowser2, pSink);
		}
		return hr;
	}

	// TODO: here should recursive to add sink to frames
	EVENT_SINK es;
	LPDISPATCH pSinkDisp = pSink->GetIDispatch(FALSE);
	BOOL br = AfxConnectionAdvise(pWebBrowser2, DIID_DWebBrowserEvents2, pSinkDisp, FALSE, &es.dwCookie);
	if (br)
	{
		es.pUnkSink = pSinkDisp;
		es.pUnkSrc = pWebBrowser2;
		es.iid = DIID_DWebBrowserEvents2;
		this->AddSink(es);
	}
	return (br) ? S_OK : E_UNEXPECTED;
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

HWND SearchWindowEx(HWND hwndParent, HWND hwndChildAfter, LPCTSTR lpszClass, LPCTSTR lpszWindow)
{
	HWND hWnd = NULL, htWnd;
	TCHAR szClassName[256];
	do {
		hWnd = FindWindowEx(hwndParent, hwndChildAfter, NULL, NULL);
		if (!hWnd) return NULL;
		GetClassName(hWnd,  szClassName,  sizeof(szClassName));
		if (_tcscmp(szClassName,  lpszClass) == 0) return hWnd;
		else {
			htWnd = SearchWindowEx(hWnd, NULL, lpszClass, lpszWindow);
			if (htWnd) return htWnd;
		}
		hwndChildAfter = hWnd;
	} while(hWnd);
	return NULL;
}

HRESULT CSpyer::AddIESink(HWND hWnd, CIESink* pSink)
{
	if (!pSink) return E_INVALIDARG;
	if (!hWnd) return AddIESink((IWebBrowser2*)NULL, pSink);
	if (!IsWindow(hWnd)) return E_INVALIDARG;

	HWND hDocViewWnd = NULL;

	EnumChildWindows(hWnd, EnumIEWindow, (LPARAM)&hDocViewWnd);
	// hDocViewWnd = SearchWindowEx(hWnd, NULL, _T("Internet Explorer_Server"), NULL);
	if (hDocViewWnd)
	{
		UINT nMsg = RegisterWindowMessage(_T("WM_HTML_GETOBJECT"));
		LRESULT lRes;
		SendMessageTimeout(hDocViewWnd, nMsg, 0L, 0L, SMTO_ABORTIFHUNG, 1000, (DWORD*)&lRes);

		CComPtr<IHTMLDocument2> spDispDoc = NULL;
		// ObjectFromLresult方法只能得到IID_IHTMLDocument2对象, 无法得到IID_IWebBrowser2对象
		// Requirements
		//  Windows NT/2000/XP/Server 2003: Included in Windows XP and Windows Server 2003.
		//  Windows 95/98/Me: Unsupported.
		//  Redistributable: Requires Active Accessibility 2.0 RDK on Windows NT 4.0 SP6 and Windows 98.
		//  Header: Declared in Oleacc.h.
		//  Library: Use Oleacc.lib.
		HRESULT hr = ObjectFromLresult (lRes, IID_IHTMLDocument2, 0 , (LPVOID*)&spDispDoc);
		if (FAILED(hr) || !spDispDoc) return hr;

		IWebBrowser2* pWebBrowser = iehelper::GetWebBrowser(spDispDoc);
		return this->AddIESink(pWebBrowser, pSink);
	}
	return E_UNEXPECTED;
}

HRESULT CSpyer::AddFlashSink(IDispatch* pDisp, CFlashSink* pSink)
{
	if (!pDisp || !pSink) return E_INVALIDARG;

	HRESULT hr = S_OK;
	EVENT_SINK es;
	hr = AtlAdvise(pDisp, pSink, DIID__IShockwaveFlashEvents, &es.dwCookie);
	if (SUCCEEDED(hr))
	{
		es.pUnkSink = pSink;
		es.pUnkSrc = pDisp;
		es.iid = DIID__IShockwaveFlashEvents;
		this->AddSink(es);
	}
	return hr;
}

HRESULT CSpyer::AddFlashSink(IWebBrowser2* pWebBrowser2, CFlashSink* pSink)
{
	if (!pWebBrowser2 || !pSink) return E_INVALIDARG;

	CComPtr<IDispatch> spDispDoc;
	HRESULT hr = pWebBrowser2->get_Document(&spDispDoc);
	if (FAILED(hr) || !spDispDoc) return hr;

	CComQIPtr<IHTMLDocument2> spDocument2 = spDispDoc;
	return AddFlashSink2Document(spDocument2, pSink);
}

HRESULT CSpyer::AddSink(EVENT_SINK eventSink)
{
	pair<EVENT_SINK_MAP::iterator, bool> retPair = 
		m_eventSinks.insert(EVENT_SINK_PAIR(eventSink.pUnkSink, eventSink));
	return (retPair.second) ? S_OK : E_OUTOFMEMORY;
}

HRESULT CSpyer::RemoveAllSinks()
{
	EVENT_SINK_MAP::iterator it;
	for(it = m_eventSinks.begin(); it != m_eventSinks.end(); ++it)
	{
		RemoveSink((*it).first, (*it).second);
	}
	m_eventSinks.clear();
	return S_OK;
}

HRESULT CSpyer::RemoveSink(LPVOID pSink)
{
	HRESULT hr = S_OK;
	EVENT_SINK_MAP::iterator it = m_eventSinks.find(pSink);
	if (it != m_eventSinks.end())
	{
		EVENT_SINK es = (*it).second;
		hr = RemoveSink(pSink, es);
	}
	return hr;
}

HRESULT CSpyer::RemoveSink(LPVOID pSink, EVENT_SINK es)
{
	HRESULT hr = S_OK;
	if(es.dwCookie > 0 && es.pUnkSink)
	{
		hr = AtlUnadvise(es.pUnkSrc, es.iid, es.dwCookie);
	}
	return hr;
}

HRESULT CSpyer::AddFlashSink2Document(IHTMLDocument2* pDocument2, CFlashSink* pSink)
{
	if (!pDocument2 || !pSink) return E_INVALIDARG;
	AddFlashSink2Embeds(pDocument2, pSink);
	AddFlashSink2Objects(pDocument2, pSink);
	AddFlashSink2Frames(pDocument2, pSink);
	return S_OK;
}

HRESULT CSpyer::AddFlashSink2Embeds(IHTMLDocument2* pDocument2, CFlashSink* pSink)
{
	if (!pDocument2 || !pSink) return E_INVALIDARG;

	CComPtr<IHTMLElementCollection> spEmbedCollection;
	HRESULT hr = pDocument2->get_embeds(&spEmbedCollection);
	if (FAILED(hr)) return hr;

	long nEmbedCount = 0;
	hr = spEmbedCollection->get_length(&nEmbedCount);
	if (FAILED(hr)) return hr;

	IDispatch* pDisp = NULL;
	CComQIPtr<IHTMLEmbedElement> spEmbedElement;
	for(long i = 0; i < nEmbedCount; i++)
	{
		hr = spEmbedCollection->item(CComVariant(i), CComVariant(), &pDisp);
		if (FAILED(hr) || !pDisp) continue;
		spEmbedElement = pDisp;
	}
	if (pDisp) 
	{
		pDisp->Release();
		pDisp = NULL;
	}
	return S_OK;
}

HRESULT CSpyer::AddFlashSink2Objects(IHTMLDocument2* pDocument2, CFlashSink* pSink)
{
	if (!pDocument2 || !pSink) return E_INVALIDARG;

	CComPtr< IHTMLElementCollection > spAllCollection;
	HRESULT hr = pDocument2->get_all(&spAllCollection);
	if (FAILED(hr)) return hr;

	IDispatch *pDisp = NULL;
	hr = spAllCollection->tags(CComVariant(_T("OBJECT")), &pDisp);
	if (FAILED(hr)) return hr;

	CComQIPtr <IHTMLElementCollection> spObjectCollection = pDisp;
	long nObjectCount = 0;
	hr = spObjectCollection->get_length(&nObjectCount);
	if (FAILED(hr)) return hr;

	CComQIPtr< IHTMLObjectElement > spObjectElement;

	for(long i = 0; i < nObjectCount; i++)
	{
		hr = spObjectCollection->item(CComVariant(i), CComVariant(), &pDisp);
		if (FAILED(hr)) continue;
		hr = spObjectElement->get_object(&pDisp);
		if (FAILED(hr) || !pDisp) continue;
		hr = AddFlashSink(pDisp, pSink);
	}
	if (pDisp) 
	{
		pDisp->Release();
		pDisp = NULL;
	}
	return S_OK;
}

HRESULT CSpyer::AddFlashSink2Frames(IHTMLDocument2* pDocument2, CFlashSink* pSink)
{
	if (!pDocument2 || !pSink) return E_INVALIDARG;

	CComPtr< IHTMLFramesCollection2 > spFramesCollection2;
	HRESULT hr = pDocument2->get_frames(&spFramesCollection2);
	if (FAILED(hr)) return hr;

	long nFrameCount = 0;
	hr = spFramesCollection2->get_length(&nFrameCount);
	if (FAILED(hr)) return hr;

	for(long i = 0; i < nFrameCount; i++)
	{
		CComVariant vDispWin2;
		hr = spFramesCollection2->item(&CComVariant(i), &vDispWin2);
		if (FAILED(hr))	continue;

		CComQIPtr<IHTMLWindow2> spWin2 = vDispWin2.pdispVal;
		if(!spWin2)	continue;

		CComQIPtr <IHTMLDocument2> spDoc2;
		spWin2->get_document(&spDoc2);

		AddFlashSink2Document(spDoc2, pSink);
	}
	return S_OK;
}

}