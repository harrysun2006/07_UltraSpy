#include "stdafx.h"
#include "FlashSink.h"

STDMETHODIMP CFlashSink::Invoke(DISPID id, REFIID riid, LCID lcid, WORD wFlags, 
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

void CFlashSink::OnReadyStateChange(long newState)
{
}

void CFlashSink::OnProgress(long percentDone)
{
}

void CFlashSink::OnFSCommand(BSTR command, BSTR args)
{
}

void CFlashSink::OnFlashCall(BSTR request)
{
}

void CFlashSink::OnGotFocus()
{
}

void CFlashSink::OnLostFocus()
{
}