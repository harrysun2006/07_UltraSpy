#include "StdAfx.h"
#include "WebMonitor.h"

using namespace std;

STDMETHODIMP CWebMonitor::QueryInterface(REFIID riid, void** ppvObject)
{
    *ppvObject = NULL;

    if(IsEqualGUID(riid, DIID_DWebBrowserEvents2) ||
        IsEqualGUID(riid, IID_IDispatch))
    {
        *ppvObject = (DWebBrowserEvents2*)this;
        AddRef();
        return S_OK;
    }
    else if(IsEqualGUID(riid, IID_IDocHostUIHandler) ||
        IsEqualGUID(riid, IID_IUnknown))
    {
        *ppvObject = (IDocHostUIHandler*)this;
        AddRef();
        return S_OK;
    }

    return E_NOINTERFACE;
}

ULONG __stdcall CWebMonitor::AddRef()
{
    m_uRefCount++;
    return m_uRefCount;
}

ULONG __stdcall CWebMonitor::Release()
{
    m_uRefCount--;
    ULONG uRefCount = m_uRefCount;
    if(uRefCount == 0)
        delete this;

    return uRefCount;
}

// IDispatch Methods
STDMETHODIMP CWebMonitor::GetTypeInfoCount(unsigned int FAR* pctinfo)
{
    return E_NOTIMPL;
}
 
STDMETHODIMP CWebMonitor::GetTypeInfo(UINT itinfo, LCID lcid, ITypeInfo** pptinfo)        
{
    return E_NOTIMPL;
}
 
 
STDMETHODIMP CWebMonitor::GetIDsOfNames(REFIID riid, OLECHAR FAR* FAR* rgszNames, 
    unsigned int cNames, LCID lcid, DISPID FAR* rgDispId)
{
    return E_NOTIMPL;
}

STDMETHODIMP CWebMonitor::Invoke(DISPID dispidMember,REFIID riid, LCID lcid, WORD wFlags,
    DISPPARAMS* pDispParams, VARIANT* pvarResult,
    EXCEPINFO* pexcepinfo, UINT* puArgErr)
{
	cout
		<< _T("CWebMonitor>>Invoke: ")
		<< _T("dispID = ") << dispidMember
		<< _T(", pParams = [");
	for (int i = 0; i < pDispParams->cArgs; i++)
	{
		cout
			<< pDispParams->rgvarg[i].bstrVal << _T(", ");
	}
	cout << "]" << endl;
    if(!pDispParams)
        return E_INVALIDARG;

    switch(dispidMember)
    {
        //
        // The parameters for this DISPID are as follows:
        // [0]: URL to navigate to - VT_BYREF|VT_VARIANT
        // [1]: An object that evaluates to the top-level or frame
        //      WebBrowser object corresponding to the event. 
    case DISPID_NAVIGATECOMPLETE2:
         
        //
        // The IDocHostUIHandler association must be set
        // up every time we navigate to a new page.
        //
        if(pDispParams->cArgs >= 2 && pDispParams->rgvarg[1].vt == VT_DISPATCH)
            SetCustomDoc(pDispParams->rgvarg[1].pdispVal);
        else
            return E_INVALIDARG;
         
        break;
         
    default:
        break;
    }
     
    return S_OK;
}
     
STDMETHODIMP CWebMonitor::ShowContextMenu(DWORD dwID, POINT FAR* ppt, IUnknown FAR* pcmdtReserved,
                          IDispatch FAR* pdispReserved)
{
    return E_NOTIMPL;
}
 
STDMETHODIMP CWebMonitor::GetHostInfo(DOCHOSTUIINFO FAR* pInfo)
{
    if(pInfo != NULL)
    {
        pInfo->dwFlags |= (m_bEnable3DBorder ? 0 : DOCHOSTUIFLAG_NO3DBORDER);
        pInfo->dwFlags |= (m_bEnableScrollBar ? 0 : DOCHOSTUIFLAG_SCROLL_NO);
    }

    return S_OK;
}
 
STDMETHODIMP CWebMonitor::ShowUI(DWORD dwID, IOleInPlaceActiveObject FAR* pActiveObject,
                IOleCommandTarget FAR* pCommandTarget,
                IOleInPlaceFrame  FAR* pFrame,
                IOleInPlaceUIWindow FAR* pDoc)
{
    return E_NOTIMPL;
}
 
STDMETHODIMP CWebMonitor::HideUI()
{
    return E_NOTIMPL;
}
 
STDMETHODIMP CWebMonitor::UpdateUI()
{
    return E_NOTIMPL;
}
 
STDMETHODIMP CWebMonitor::EnableModeless(BOOL fEnable)
{
    return E_NOTIMPL;
}

STDMETHODIMP CWebMonitor::OnDocWindowActivate(BOOL fActivate)
{
    return E_NOTIMPL;
}

STDMETHODIMP CWebMonitor::OnFrameWindowActivate(BOOL fActivate)
{
    return E_NOTIMPL;
}

STDMETHODIMP CWebMonitor::ResizeBorder(LPCRECT prcBorder, IOleInPlaceUIWindow FAR* pUIWindow,
                       BOOL fRameWindow)
{
    return E_NOTIMPL;
}

STDMETHODIMP CWebMonitor::TranslateAccelerator(LPMSG lpMsg, const GUID FAR* pguidCmdGroup,
                               DWORD nCmdID)
{
    return E_NOTIMPL;
}

STDMETHODIMP CWebMonitor::GetOptionKeyPath(LPOLESTR FAR* pchKey, DWORD dw)
{
    return E_NOTIMPL;
}

STDMETHODIMP CWebMonitor::GetDropTarget(IDropTarget* pDropTarget,
                        IDropTarget** ppDropTarget)
{
    return E_NOTIMPL;
}

STDMETHODIMP CWebMonitor::GetExternal(IDispatch** ppDispatch)
{
    return E_NOTIMPL;
}

STDMETHODIMP CWebMonitor::TranslateUrl(DWORD dwTranslate, OLECHAR* pchURLIn,
                       OLECHAR** ppchURLOut)
{
    return E_NOTIMPL;
}

STDMETHODIMP CWebMonitor::FilterDataObject(IDataObject* pDO, IDataObject** ppDORet)
{
    return E_NOTIMPL;
}
 
HRESULT CWebMonitor::SetWebBrowser(IWebBrowser2* pWebBrowser2)
{
    // Unadvise the event sink, if there was a
    // previous reference to the WebBrowser control.
    if(m_pWebBrowser2)
    {
		return S_OK;
        _CoUnadvise(m_pWebBrowser2, DIID_DWebBrowserEvents2, m_dwCookie);
        m_dwCookie = 0;

        m_pWebBrowser2->Release();
    }

    m_pWebBrowser2 = pWebBrowser2;
    if(pWebBrowser2 == NULL)
        return S_OK;

    m_pWebBrowser2->AddRef();

    return _CoAdvise(m_pWebBrowser2, (IDispatch*)this, DIID_DWebBrowserEvents2, &m_dwCookie);
}

void CWebMonitor::Enable3DBorder(BOOL bEnable)
{
    m_bEnable3DBorder = bEnable;
}

void CWebMonitor::EnableScrollBar(BOOL bEnable)
{
    m_bEnableScrollBar = bEnable;
}

void CWebMonitor::SetCustomDoc(LPDISPATCH lpDisp)
{
    if(lpDisp == NULL)
        return;

    IWebBrowser2* pWebBrowser2 = NULL;
    HRESULT hr = lpDisp->QueryInterface(IID_IWebBrowser2, (void**)&pWebBrowser2);

    if(SUCCEEDED(hr) && pWebBrowser2)
    {
        IDispatch* pDoc = NULL;
        hr = pWebBrowser2->get_Document(&pDoc);
     
        if(SUCCEEDED(hr) && pDoc)
        {
            ICustomDoc* pCustDoc = NULL;
            hr = pDoc->QueryInterface(IID_ICustomDoc, (void**)&pCustDoc);
            if(SUCCEEDED(hr) && pCustDoc != NULL)
            {
                   pCustDoc->SetUIHandler(this);
                   pCustDoc->Release();
            }

            pDoc->Release();
        }

        pWebBrowser2->Release();
    }
}
