// Spyer.h: interface for the CSpyer class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(SPYER__INCLUDED_)
#define SPYER__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "resource.h"
#include "IESink.h"
#include "FlashSink.h"
#include <mshtml.h>
#include <map>

using namespace std;

namespace spyer {
typedef struct EVENT_SINK {
	LPUNKNOWN pUnkSrc;
	LPUNKNOWN pUnkSink;
	DWORD dwCookie;
	IID iid;
} EVENT_SINK;

typedef EVENT_SINK* LPEVENT_SINK;
typedef map<LPVOID, EVENT_SINK> EVENT_SINK_MAP;
typedef pair<LPVOID, EVENT_SINK> EVENT_SINK_PAIR;

using namespace ShockwaveFlashObjects;

class CSpyer
{
public:
	static CSpyer* Instance()
	{
		if(_instance == NULL) _instance = new CSpyer();
		return _instance;
	}
	HRESULT AddIESink(HWND hWnd, CIESink* pSink);
	HRESULT AddIESink(IWebBrowser2* pWebBrowser2, CIESink* pSink);
	HRESULT AddFlashSink(IWebBrowser2* pWebBrowser2, CFlashSink* pSink);
	HRESULT AddFlashSink(IDispatch* pDisp, CFlashSink* pSink);
	HRESULT RemoveSink(LPVOID pSink);
private:
	CSpyer(void){};
	virtual ~CSpyer(void){RemoveAllSinks();};
	static CSpyer* _instance;
	EVENT_SINK_MAP m_eventSinks;
	HRESULT AddSink(EVENT_SINK eventSink);
	HRESULT RemoveAllSinks();
	HRESULT RemoveSink(LPVOID pSink, EVENT_SINK es);
	HRESULT AddFlashSink2Document(IHTMLDocument2* pDocument2, CFlashSink* pSink);
	HRESULT AddFlashSink2Embeds(IHTMLDocument2* pDocument2, CFlashSink* pSink);
	HRESULT AddFlashSink2Objects(IHTMLDocument2* pDocument2, CFlashSink* pSink);
	HRESULT AddFlashSink2Frames(IHTMLDocument2* pDocument2, CFlashSink* pSink);
};
}
#endif // !defined(SPYER__INCLUDED_)