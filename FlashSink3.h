// FlashSink3.h: interface for the CFlashSink3 class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(FLASHSINK3__INCLUDED_)
#define FLASHSINK3__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "resource.h"
#include "Flash9d.tlh"

#define DISPID_ONREADYSTATECHANGE	-609//0xFFFFFD9F
#define DISPID_ONPROGRESS			0x000007A6
#define DISPID_FSCOMMAND			0x00000096
#define DISPID_FLASHCALL			0x000000C5
#define DISPID_GOTFOCUS				0x800100E0
#define DISPID_LOSTFOCUS			0x800100E1

class CFlashSink3 : public IDispatchImpl<ShockwaveFlashObjects::_IShockwaveFlashEvents, 
	&ShockwaveFlashObjects::DIID__IShockwaveFlashEvents,
	&ShockwaveFlashObjects::LIBID_ShockwaveFlashObjects>
//class CFlashSink3 : public IShockwaveFlashEvents
{
public:
	STDMETHOD(QueryInterface)(const struct _GUID &iid,void ** ppv);
	virtual ULONG STDMETHODCALLTYPE AddRef(void);
	virtual ULONG STDMETHODCALLTYPE Release(void);
	STDMETHOD(Invoke)(DISPID, REFIID, LCID, WORD, DISPPARAMS*,VARIANT*, EXCEPINFO*, UINT*);
protected:
	virtual void OnReadyStateChange(long newState);
	virtual void OnProgress(long percentDone);
	virtual void OnFSCommand(BSTR command, BSTR args);
	virtual void OnFlashCall(BSTR request);
	virtual void OnGotFocus();
	virtual void OnLostFocus();
};

#endif // !defined(FLASHSINK3__INCLUDED_)