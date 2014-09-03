#include "stdafx.h"
#include "IESink2.h"
#include <fstream>
#include <exdispid.h>
#include <mshtml.h>

using namespace std;

void CIESink2::EnumDocument(IHTMLDocument2 *pIHTMLDocument2)
{
	HRESULT hr;
	USES_CONVERSION;

	if (!pIHTMLDocument2) return;

	CComBSTR bstrTitle, bstrURL, bstrHTML;
	pIHTMLDocument2->get_title(&bstrTitle);	//取得文档标题
	pIHTMLDocument2->get_URL(&bstrURL);

	CComQIPtr <IHTMLElement> spBody;
	hr = pIHTMLDocument2->get_body(&spBody);
	if (FAILED(hr)) return;
	
	CComQIPtr <IHTMLElement> spHTML;
	hr = spBody->get_parentElement(&spHTML);
	if (FAILED(hr)) return;
	spHTML->get_outerHTML(&bstrHTML);
	
	cout << _T("CIESink2::OnDocumentComplete.Title: ") << (bstrTitle ? OLE2CT(bstrTitle) : _T("")) << endl;
	cout << _T("CIESink2::OnDocumentComplete.URL: ") << (bstrURL ? OLE2CT(bstrURL) : _T("")) << endl;
}

STDMETHODIMP CIESink2::Invoke(DISPID id, REFIID riid, LCID lcid, WORD wFlags, 
	DISPPARAMS* pdp, VARIANT* pvarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr)
{
	REDIR_OUT2LOG;
	USES_CONVERSION;
	if (!pdp) return E_INVALIDARG;

	switch (id)
	{
		//
		// The parameters for this DISPID are as follows:
		// [0]: Cancel flag  - VT_BYREF|VT_BOOL
		// [1]: HTTP headers - VT_BYREF|VT_VARIANT
		// [2]: Address of HTTP POST data  - VT_BYREF|VT_VARIANT 
		// [3]: Target frame name - VT_BYREF|VT_VARIANT 
		// [4]: Option flags - VT_BYREF|VT_VARIANT
		// [5]: URL to navigate to - VT_BYREF|VT_VARIANT
		// [6]: An object that evaluates to the top-level or frame
		//      WebBrowser object corresponding to the event. 
		//
		case DISPID_BEFORENAVIGATE2:
			ATLASSERT(pdp->cArgs == 7);
			OnBeforeNavigate2(pdp->rgvarg[6].pdispVal,
							  pdp->rgvarg[5].pvarVal,
							  pdp->rgvarg[4].pvarVal,
							  pdp->rgvarg[3].pvarVal,
							  pdp->rgvarg[2].pvarVal,
							  pdp->rgvarg[1].pvarVal,
							  pdp->rgvarg[0].pboolVal);
		break;

		//
		// The parameters for this DISPID:
		// [0]: URL navigated to - VT_BYREF|VT_VARIANT
		// [1]: An object that evaluates to the top-level or frame
		//      WebBrowser object corresponding to the event. 
		//
		case DISPID_NAVIGATECOMPLETE2:
			ATLASSERT(pdp->cArgs == 2);
			OnNavigateComplete2(pdp->rgvarg[1].pdispVal,
								pdp->rgvarg[0].pvarVal);

		break;

		//
		// The parameters for this DISPID:
		// [0]: New status bar text - VT_BSTR
		//
		case DISPID_STATUSTEXTCHANGE:
			ATLASSERT(pdp->cArgs == 1);
			OnStatusTextChange(W2CT(pdp->rgvarg[0].bstrVal));
		break;

		//
		// The parameters for this DISPID:
		// [0]: Maximum progress - VT_I4
		// [1]: Amount of total progress - VT_I4
		//
		case DISPID_PROGRESSCHANGE:
			ATLASSERT(pdp->cArgs == 2);
			OnProgressChange(pdp->rgvarg[1].lVal,
							 pdp->rgvarg[0].lVal);
		break;

		//
		// The parameters for this DISPID:
		// [0]: URL navigated to - VT_BYREF|VT_VARIANT
		// [1]: An object that evaluates to the top-level or frame
		//      WebBrowser object corresponding to the event. 
		//
		case DISPID_DOCUMENTCOMPLETE:
			ATLASSERT(pdp->cArgs == 2);
			OnDocumentComplete(pdp->rgvarg[1].pdispVal,
							   pdp->rgvarg[0].pvarVal);
		break;

		case DISPID_DOWNLOADBEGIN:
			ATLASSERT(pdp->cArgs == 0);
			OnDownloadBegin();
		break;

		case DISPID_DOWNLOADCOMPLETE:
			ATLASSERT(pdp->cArgs == 0);
			OnDownloadComplete();
		break;

		//
		// The parameters for this DISPID:
		// [0]: Enabled state - VT_BOOL
		// [1]: Command identifier - VT_I4
		//
		case DISPID_COMMANDSTATECHANGE:
			ATLASSERT(pdp->cArgs == 2);
			OnCommandStateChange(pdp->rgvarg[1].lVal,
								 pdp->rgvarg[0].boolVal);
		break;

		//
		// The parameters for this DISPID:
		// [0]: Cancel flag  - VT_BYREF|VT_BOOL
		// [1]: An object that evaluates to the top-level or frame
		//      WebBrowser object corresponding to the event. 
		//
		case DISPID_NEWWINDOW2:
			ATLASSERT(pdp->cArgs == 2);
			OnNewWindow2(pdp->rgvarg[1].ppdispVal,
						 pdp->rgvarg[0].pboolVal);
		break;

		//
		// The parameters for this DISPID:
		// [0]: URL to navigate to - VT_BSTR
		// [1]: URL Context - VT_BSTR
		// [2]: dwFlag - VT_ULONG
		// [3]: Cancel flag  - VT_BYREF|VT_BOOL
		// [4]: An object that evaluates to the top-level or frame
		//      WebBrowser object corresponding to the event. 
		//
		case DISPID_NEWWINDOW3:
			ATLASSERT(pdp->cArgs == 5);
			OnNewWindow3(pdp->rgvarg[4].ppdispVal,
						 pdp->rgvarg[3].pboolVal,
						 pdp->rgvarg[2].ulVal,
						 pdp->rgvarg[1].bstrVal,
						 pdp->rgvarg[0].bstrVal);
		break;

		//
		// The parameters for this DISPID:
		// [0]: Document title - VT_BSTR
		//
		case DISPID_TITLECHANGE:
			ATLASSERT(pdp->cArgs == 1);
			OnTitleChange(W2CT(pdp->rgvarg[0].bstrVal));
		break;

		//
		// The parameters for this DISPID:
		// [0]: Name of property that changed - VT_BSTR
		//
		case DISPID_PROPERTYCHANGE:
			ATLASSERT(pdp->cArgs == 1);
			OnPropertyChange(W2CT(pdp->rgvarg[0].bstrVal));
		break;

		//
		// The parameters for this DISPID:
		// [0]: Address of cancel flag - VT_BYREF|VT_BOOL
		//
		case DISPID_QUIT:
		case DISPID_ONQUIT:
			ATLASSERT(pdp->cArgs == 0);
			OnQuit();
		break;

		//
		// The parameters for this DISPID:
		// [0]: Browser shown/hidden flag - VT_BOOL
		//	
		case DISPID_ONVISIBLE:
			ATLASSERT(pdp->cArgs == 1);
			OnVisible(pdp->rgvarg[0].boolVal);
		break;	

		//
		// The parameters for this DISPID:
		// [0]: Toolbar shown/hidden flag - VT_BOOL
		//	
		case DISPID_ONTOOLBAR:
			ATLASSERT(pdp->cArgs == 1);
			OnToolBar(pdp->rgvarg[0].boolVal);
		break;

		//
		// The parameters for this DISPID:
		// [0]: Menubar shown/hidden flag - VT_BOOL
		//	
		case DISPID_ONMENUBAR:
			ATLASSERT(pdp->cArgs == 1);
			OnMenuBar(pdp->rgvarg[0].boolVal);
		break;

		//
		// The parameters for this DISPID:
		// [0]: Statusbar shown/hidden flag - VT_BOOL
		//	
		case DISPID_ONSTATUSBAR:
			ATLASSERT(pdp->cArgs == 1);
			OnStatusBar(pdp->rgvarg[0].boolVal);
		break;

		//
		// The parameters for this DISPID:
		// [0]: Fullscreen mode on/off flag - VT_BOOL
		//	
		case DISPID_ONFULLSCREEN:
			ATLASSERT(pdp->cArgs == 1);
			OnFullScreen(pdp->rgvarg[0].boolVal);
		break;

		//
		// The parameters for this DISPID:
		// [0]: Theater mode on/off flag - VT_BOOL
		//	
		case DISPID_ONTHEATERMODE:
			ATLASSERT(pdp->cArgs == 1);
			OnTheaterMode(pdp->rgvarg[0].boolVal);
		break;

		//
		// The parameters for this DISPID:
		// [0]: Host window allow/disallow resizing flag - VT_BOOL
		//	
		case DISPID_WINDOWSETRESIZABLE:
			ATLASSERT(pdp->cArgs == 1);
			OnWindowSetResizable(pdp->rgvarg[0].boolVal);
		break;

		//
		// The parameters for this DISPID:
		// [0]: The new left coordinate of the host window change to - VT_I4
		//	
		case DISPID_WINDOWSETLEFT:
			ATLASSERT(pdp->cArgs == 1);
			OnWindowSetLeft(pdp->rgvarg[0].lVal);
		break;

		//
		// The parameters for this DISPID:
		// [0]: The new top coordinate of the host window change to - VT_I4
		//	
		case DISPID_WINDOWSETTOP:
			ATLASSERT(pdp->cArgs == 1);
			OnWindowSetTop(pdp->rgvarg[0].lVal);
		break;

		//
		// The parameters for this DISPID:
		// [0]: The new width of the host window change to - VT_I4
		//	
		case DISPID_WINDOWSETWIDTH:
			ATLASSERT(pdp->cArgs == 1);
			OnWindowSetWidth(pdp->rgvarg[0].lVal);
		break;

		//
		// The parameters for this DISPID:
		// [0]: The new height of the host window change to - VT_I4
		//	
		case DISPID_WINDOWSETHEIGHT:
			ATLASSERT(pdp->cArgs == 1);
			OnWindowSetHeight(pdp->rgvarg[0].lVal);
		break;

		//
		// The parameters for this DISPID:
		// [0]: Cancel flag  - VT_BYREF|VT_BOOL
		// [1]: Is a child window to close me - VT_BOOL
		//	
		case DISPID_WINDOWCLOSING:
			ATLASSERT(pdp->cArgs == 2);
			OnWindowClosing(pdp->rgvarg[1].boolVal, pdp->rgvarg[0].pboolVal);
		break;

		case DISPID_CLIENTTOHOSTWINDOW:
			ATLASSERT(pdp->cArgs == 2);
			OnClientToHostWindow(pdp->rgvarg[1].plVal, pdp->rgvarg[0].plVal);
		break;

		case DISPID_SETSECURELOCKICON:
			ATLASSERT(pdp->cArgs == 1);
			OnSetSecureLockIcon(pdp->rgvarg[0].lVal);
		break;
/*
		case DISPID_FILEDOWNLOAD:
			ATLASSERT(pdp->cArgs == 1);
			OnFileDownload(pdp->rgvarg[0].pboolVal);
		break;
*/
		case DISPID_NAVIGATEERROR:
			ATLASSERT(pdp->cArgs == 5);
			OnNavigateError(pdp->rgvarg[4].pdispVal, 
							pdp->rgvarg[3].pvarVal,
							pdp->rgvarg[2].pvarVal,
							pdp->rgvarg[1].pvarVal,
							pdp->rgvarg[0].pboolVal);
		break;
		
		case DISPID_PRINTTEMPLATEINSTANTIATION:
			ATLASSERT(pdp->cArgs == 1);
			OnPrintTemplateInstantiation(pdp->rgvarg[0].pdispVal);
		break;
		
		case DISPID_PRINTTEMPLATETEARDOWN:
			ATLASSERT(pdp->cArgs == 1);
			OnPrintTemplateTeardown(pdp->rgvarg[0].pdispVal);
		break;
		
		case DISPID_UPDATEPAGESTATUS:
			ATLASSERT(pdp->cArgs == 3);
			OnUpdatePageStatus(pdp->rgvarg[2].pdispVal,
							   pdp->rgvarg[1].pvarVal,
							   pdp->rgvarg[0].pvarVal);
		break;

		case DISPID_PRIVACYIMPACTEDSTATECHANGE:
			ATLASSERT(pdp->cArgs == 1);
			OnPrivacyImpactedStateChange(pdp->rgvarg[0].boolVal);
		break;

		default:
			cout
				<< _T("CIESink2::Invoke unknown event: ")
				<< _T("dispID = ") << id
				<< _T(", pParams = [");
			for (int i = 0; i < pdp->cArgs; i++)
			{
				cout << pdp->rgvarg[i].bstrVal << _T(", ");
			}
			cout << "]" << endl;
		break;
	}
	REDIR_OUT2OUT;
	return S_OK;
}

// DWebBrowserEvents2 methods
// Fired before navigate occurs in the given WebBrowser (window or frameset element). 
// The processing of this navigation may be modified.
void CIESink2::OnBeforeNavigate2(LPDISPATCH pDisp,
			 				   VARIANT* URL,
							   VARIANT* Flags,
							   VARIANT* TargetFrameName,
							   VARIANT* PostData,
							   VARIANT* Headers,
							   VARIANT_BOOL* Cancel)
{
	cout << "CIESink2::OnBeforeNavigate2" << endl;
}

// Fired when the document being navigated to becomes visible and enters the navigation stack.
void CIESink2::OnNavigateComplete2(LPDISPATCH pDisp, VARIANT* URL)
{
	cout << "CIESink2::OnNavigateComplete" << endl;
}

// Statusbar text changed.
void CIESink2::OnStatusTextChange(LPCTSTR Text)
{
	cout << "CIESink2::OnStatusTextChange" << endl;
}

// Fired when download progress is updated.
void CIESink2::OnProgressChange(long Progress, long ProgressMax)
{
	cout << "CIESink2::OnProgressChange" << endl;
}

// Fired when the document being navigated to reaches ReadyState_Complete.
void CIESink2::OnDocumentComplete(LPDISPATCH pDisp, VARIANT *pvarURL)
{
	cout << "CIESink2::OnDocumentComplete" << endl;
	CComQIPtr<IWebBrowser2> pWebBrowser2 = pDisp;
	HRESULT hr = S_FALSE;
	if(pWebBrowser2)
	{
		CComPtr<IDispatch> spDispDoc;
		hr = pWebBrowser2->get_Document(&spDispDoc);
		if (SUCCEEDED(hr) && spDispDoc) 
		{
			CComQIPtr<IHTMLDocument2> spDocument = spDispDoc;
			EnumDocument(spDocument);
		}
	}
}

// Download of a page started.
void CIESink2::OnDownloadBegin()
{
	cout << "CIESink2::OnDownloadBegin" << endl;
}

// Download of page complete.
void CIESink2::OnDownloadComplete()
{
	cout << "CIESink2::OnDownloadComplete" << endl;
}

// The enabled state of a command changed.
void CIESink2::OnCommandStateChange(long Command, BOOL Enable)
{
	cout << "CIESink2::OnCommandStateChange" << endl;
}

// A new, hidden, non-navigated WebBrowser window is needed.
void CIESink2::OnNewWindow2(LPDISPATCH* ppDisp, VARIANT_BOOL* Cancel)
{
	cout << "CIESink2::OnNewWindow2" << endl;
}

// A new, hidden, non-navigated WebBrowser window is needed.
void CIESink2::OnNewWindow3(LPDISPATCH* ppDisp, VARIANT_BOOL* Cancel, unsigned long dwFlags, 
	BSTR bstrUrlContext, BSTR bstrUrl)
{
	cout << "CIESink2::OnNewWindow3" << endl;
}

// Document title changed.
void CIESink2::OnTitleChange(LPCTSTR Text)
{
	cout << "CIESink2::OnTitleChange" << endl;
}

// Fired when the PutProperty method has been called.
void CIESink2::OnPropertyChange(LPCTSTR szProperty)
{
	cout << "CIESink2::OnPropertyChange" << endl;
}

// Fired when application is quiting.
void CIESink2::OnQuit()
{
	cout << "CIESink2::OnQuit" << endl;
}

// Fired when the window should be shown/hidden.
void CIESink2::OnVisible(VARIANT_BOOL Visible)
{
	cout << "CIESink2::OnVisible" << endl;
}

// Fired when the toolbar  should be shown/hidden.
void CIESink2::OnToolBar(VARIANT_BOOL ToolBar)
{
	cout << "CIESink2::OnToolBar" << endl;
}

// Fired when the menubar should be shown/hidden.
void CIESink2::OnMenuBar(VARIANT_BOOL MenuBar)
{
	cout << "CIESink2::OnMenuBar" << endl;
}

// Fired when the statusbar should be shown/hidden.
void CIESink2::OnStatusBar(VARIANT_BOOL StatusBar)
{
	cout << "CIESink2::OnStatusBar" << endl;
}

// Fired when fullscreen mode should be on/off.
void CIESink2::OnFullScreen(VARIANT_BOOL FullScreen)
{
	cout << "CIESink2::OnFullScreen" << endl;
}

// Fired when theater mode should be on/off.
void CIESink2::OnTheaterMode(VARIANT_BOOL TheaterMode)
{
	cout << "CIESink2::OnTheaterMode" << endl;
}

// Fired when the host window should allow/disallow resizing.
void CIESink2::OnWindowSetResizable(VARIANT_BOOL Resizable)
{
	cout << "CIESink2::OnWindowSetResizable" << endl;
}

// Fired when the host window should change its Left coordinate.
void CIESink2::OnWindowSetLeft(long Left)
{
	cout << "CIESink2::OnWindowSetLeft" << endl;
}

// Fired when the host window should change its Top coordinate.
void CIESink2::OnWindowSetTop(long Top)
{
	cout << "CIESink2::OnWindowSetTop" << endl;
}

// Fired when the host window should change its width.
void CIESink2::OnWindowSetWidth(long Width)
{
	cout << "CIESink2::OnWindowSetWidth" << endl;
}

// Fired when the host window should change its height.
void CIESink2::OnWindowSetHeight(long Height)
{
	cout << "CIESink2::OnWindowSetHeight" << endl;
}

// Fired when the WebBrowser is about to be closed by script.
void CIESink2::OnWindowClosing(VARIANT_BOOL IsChildWindow, VARIANT_BOOL* Cancel)
{
	cout << "CIESink2::OnWindowClosing" << endl;
}

// Fired to request client sizes be converted to host window sizes.
void CIESink2::OnClientToHostWindow(long* CX, long* CY)
{
	cout << "CIESink2::OnClientToHostWindow" << endl;
}

// Fired to indicate the security level of the current web page contents.
void CIESink2::OnSetSecureLockIcon(long SecureLockIcon)
{
	cout << "CIESink2::OnSetSecureLockIcon" << endl;
}

// Fired to indicate the File Download dialog is opening.
void CIESink2::OnFileDownload(VARIANT_BOOL* Cancel)
{
	cout << "CIESink2::OnFileDownload" << endl;
}

// Fired when a binding error occurs (window or frameset element).
void CIESink2::OnNavigateError(LPDISPATCH pDisp, VARIANT* URL, VARIANT* Frame, 
	VARIANT* StatusCode, VARIANT_BOOL* Cancel)
{
	cout << "CIESink2::OnNavigateError" << endl;
}

// Fired when a print template is instantiated.
void CIESink2::OnPrintTemplateInstantiation(LPDISPATCH pDisp)
{
	cout << "CIESink2::OnPrintTemplateInstantiation" << endl;
}

// Fired when a print template destroyed."
void CIESink2::OnPrintTemplateTeardown(LPDISPATCH pDisp)
{
	cout << "CIESink2::OnPrintTemplateTeardown" << endl;
}

// Fired when a page is spooled. When it is fired can be changed by a custom template.
void CIESink2::OnUpdatePageStatus(LPDISPATCH pDisp, VARIANT* nPage, VARIANT* fDone)
{
	cout << "CIESink2::OnUpdatePageStatus" << endl;
}

// Fired when the global privacy impacted state changes.
void CIESink2::OnPrivacyImpactedStateChange(VARIANT_BOOL bImpacted)
{
	cout << "CIESink2::OnPrivacyImpactedStateChange" << endl;
}
