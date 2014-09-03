#ifndef __WEBMONITOR_H__
#define __WEBMONITOR_H__

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "resource.h"

#include <MSHTMHST.H>
#include <EXDISP.H>
#include <EXDISPID.H>

inline 
HRESULT _CoAdvise(IUnknown* pUnkCP, IUnknown* pUnk, const IID& iid, LPDWORD pdw)
{
	IConnectionPointContainer* pCPC = NULL;
	IConnectionPoint* pCP = NULL;
	HRESULT hRes = pUnkCP->QueryInterface(IID_IConnectionPointContainer, (void**)&pCPC);
	if(SUCCEEDED(hRes) && pCPC != NULL)
	{
		hRes = pCPC->FindConnectionPoint(iid, &pCP);
		if(SUCCEEDED(hRes) && pCP != NULL)
		{
			hRes = pCP->Advise(pUnk, pdw);
			pCP->Release();
		}

		pCPC->Release();
	}

	return hRes;
}

inline 
HRESULT _CoUnadvise(IUnknown* pUnkCP, const IID& iid, DWORD dw)
{
    IConnectionPointContainer* pCPC = NULL;
    IConnectionPoint* pCP = NULL;
    HRESULT hRes = pUnkCP->QueryInterface(IID_IConnectionPointContainer, (void**)&pCPC);
    if(SUCCEEDED(hRes) && pCPC != NULL)
    {
        hRes = pCPC->FindConnectionPoint(iid, &pCP);
        if(SUCCEEDED(hRes) && pCP != NULL)
        {
            hRes = pCP->Unadvise(dw);
            pCP->Release();
        }
 
        pCPC->Release();
    }
 
    return hRes;
}

class CWebMonitor : public DWebBrowserEvents2, public IDocHostUIHandler
{
    ULONG m_uRefCount;
 
    IWebBrowser2* m_pWebBrowser2;
    DWORD m_dwCookie;
 
    BOOL m_bEnable3DBorder;
    BOOL m_bEnableScrollBar;
 
public:
    CWebMonitor() : m_uRefCount(0), m_pWebBrowser2(NULL), m_dwCookie(0)
    {
        m_bEnable3DBorder = TRUE;
        m_bEnableScrollBar = TRUE;
    }
 
    virtual ~CWebMonitor()
    {
    }

protected:
    // IUnknown Methods
    STDMETHOD(QueryInterface)(REFIID riid, void** ppvObject);
    STDMETHOD_(ULONG, AddRef)(void);
    STDMETHOD_(ULONG, Release)(void);
    // IDispatch Methods
    STDMETHOD(GetTypeInfoCount)(unsigned int FAR* pctinfo);
    STDMETHOD(GetTypeInfo)(UINT itinfo, LCID lcid, ITypeInfo** pptinfo);        
    STDMETHOD(GetIDsOfNames)(REFIID riid, OLECHAR FAR* FAR* rgszNames, 
        unsigned int cNames, LCID lcid, DISPID FAR* rgDispId);
    STDMETHOD(Invoke)(DISPID dispidMember,REFIID riid, LCID lcid, WORD wFlags,
        DISPPARAMS* pDispParams, VARIANT* pvarResult,
        EXCEPINFO* pexcepinfo, UINT* puArgErr);
     
    // IDocHostUIHandler Methods
protected:
    STDMETHOD(ShowContextMenu)(DWORD dwID, POINT FAR* ppt, IUnknown FAR* pcmdtReserved,
                              IDispatch FAR* pdispReserved);
    STDMETHOD(GetHostInfo)(DOCHOSTUIINFO FAR* pInfo);
    STDMETHOD(ShowUI)(DWORD dwID, IOleInPlaceActiveObject FAR* pActiveObject,
                    IOleCommandTarget FAR* pCommandTarget,
                    IOleInPlaceFrame  FAR* pFrame,
                    IOleInPlaceUIWindow FAR* pDoc);
    STDMETHOD(HideUI)(void);
    STDMETHOD(UpdateUI)(void);
    STDMETHOD(EnableModeless)(BOOL fEnable);
    STDMETHOD(OnDocWindowActivate)(BOOL fActivate);
    STDMETHOD(OnFrameWindowActivate)(BOOL fActivate);
    STDMETHOD(ResizeBorder)(LPCRECT prcBorder, IOleInPlaceUIWindow FAR* pUIWindow,
                           BOOL fRameWindow);
    STDMETHOD(TranslateAccelerator)(LPMSG lpMsg, const GUID FAR* pguidCmdGroup,
                                   DWORD nCmdID);
    STDMETHOD(GetOptionKeyPath)(LPOLESTR FAR* pchKey, DWORD dw);
    STDMETHOD(GetDropTarget)(IDropTarget* pDropTarget,
                            IDropTarget** ppDropTarget);
    STDMETHOD(GetExternal)(IDispatch** ppDispatch);
    STDMETHOD(TranslateUrl)(DWORD dwTranslate, OLECHAR* pchURLIn,
                           OLECHAR** ppchURLOut);
    STDMETHOD(FilterDataObject)(IDataObject* pDO, IDataObject** ppDORet);
 
public:
    HRESULT SetWebBrowser(IWebBrowser2* pWebBrowser2);
    void Enable3DBorder(BOOL bEnable = TRUE);
    void EnableScrollBar(BOOL bEnable = TRUE);

private:
    void SetCustomDoc(LPDISPATCH lpDisp);
};
 
#endif // __WEBMONITOR_H__