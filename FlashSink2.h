// FlashSink2.h: interface for the CFlashSink2 class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(FLASHSINK2__INCLUDED_)
#define FLASHSINK2__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "resource.h"

#define DISPID_ONREADYSTATECHANGE	-609
#define DISPID_ONPROGRESS			0x000007A6
#define DISPID_FSCOMMAND			0x00000096
#define DISPID_FLASHCALL			0x000000C5
#define DISPID_GOTFOCUS				0x800100E0
#define DISPID_LOSTFOCUS			0x800100E1

class CFlashSink2 : public CCmdTarget
{
public:
	CFlashSink2() {this->EnableAutomation();}
protected:
	virtual void OnReadyStateChange(long newState);
	virtual void OnProgress(long percentDone);
	virtual void OnFSCommand(BSTR command, BSTR args);
	virtual void OnFlashCall(BSTR request);
	virtual void OnGotFocus();
	virtual void OnLostFocus();
DECLARE_DISPATCH_MAP()
};

#endif // !defined(FLASHSINK2__INCLUDED_)