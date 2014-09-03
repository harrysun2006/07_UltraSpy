#include "stdafx.h"
#include "FlashSink3.h"
#include <afxctl.h>

STDMETHODIMP CFlashSink3::QueryInterface(const struct _GUID &iid,void ** ppv)
{
	*ppv = this;
	return S_OK;
}

ULONG CFlashSink3::AddRef(void)
{
	return S_OK;
}

ULONG CFlashSink3::Release(void)
{
	return S_OK;
}

STDMETHODIMP CFlashSink3::Invoke(DISPID id, REFIID riid, LCID lcid, WORD wFlags, 
	DISPPARAMS* pdp, VARIANT* pvarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr)
{
	if (!pdp) return E_INVALIDARG;

	switch (id)
	{
		case DISPID_ONREADYSTATECHANGE:
			ATLASSERT(pdp->cArgs == 1);
			OnReadyStateChange(pdp->rgvarg[0].lVal);
		break;

		case DISPID_ONPROGRESS:
			ATLASSERT(pdp->cArgs == 1);
			OnProgress(pdp->rgvarg[0].lVal);
		break;

		case DISPID_FSCOMMAND:
			ATLASSERT(pdp->cArgs == 2);
			OnFSCommand(pdp->rgvarg[1].bstrVal, pdp->rgvarg[0].bstrVal);
		break;

		case DISPID_FLASHCALL:
			ATLASSERT(pdp->cArgs == 1);
			OnFlashCall(pdp->rgvarg[0].bstrVal);
		break;

		case DISPID_GOTFOCUS:
			ATLASSERT(pdp->cArgs == 0);
			OnGotFocus();
		break;

		case DISPID_LOSTFOCUS:
			ATLASSERT(pdp->cArgs == 0);
			OnLostFocus();
		break;

		default:
		break;
	}	
	return S_OK;
}

void CFlashSink3::OnReadyStateChange(long newState)
{
}

void CFlashSink3::OnProgress(long percentDone)
{
}

void CFlashSink3::OnFSCommand(BSTR command, BSTR args)
{
}

void CFlashSink3::OnFlashCall(BSTR request)
{
}

void CFlashSink3::OnGotFocus()
{
}

void CFlashSink3::OnLostFocus()
{
}