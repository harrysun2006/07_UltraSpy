// FlashSink.h: interface for the CFlashSink class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(FLASHSINK__INCLUDED_)
#define FLASHSINK__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "resource.h"
#include "Flash9d.tlh"

#define DISPID_ONREADYSTATECHANGE	-609
#define DISPID_ONPROGRESS			0x000007A6
#define DISPID_FSCOMMAND			0x00000096
#define DISPID_FLASHCALL			0x000000C5
#define DISPID_GOTFOCUS				0x800100E0
#define DISPID_LOSTFOCUS			0x800100E1

class ATL_NO_VTABLE CFlashSink :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CFlashSink>,
	public IDispatchImpl<ShockwaveFlashObjects::_IShockwaveFlashEvents, &ShockwaveFlashObjects::DIID__IShockwaveFlashEvents,&ShockwaveFlashObjects::LIBID_ShockwaveFlashObjects>
{
public:
	STDMETHOD(Invoke)(DISPID, REFIID, LCID, WORD, DISPPARAMS*,VARIANT*, EXCEPINFO*, UINT*);
// virtual HRESULT __stdcall raw_OnReadyStateChange (long newState) = 0;

private:
	virtual void OnReadyStateChange(long newState);
	virtual void OnProgress(long percentDone);
	virtual void OnFSCommand(BSTR command, BSTR args);
	virtual void OnFlashCall(BSTR request);
	virtual void OnGotFocus();
	virtual void OnLostFocus();
BEGIN_COM_MAP(CFlashSink)
COM_INTERFACE_ENTRY(IDispatch)
COM_INTERFACE_ENTRY(ShockwaveFlashObjects::_IShockwaveFlashEvents)
END_COM_MAP()
};

#endif // !defined(FLASHSINK__INCLUDED_)