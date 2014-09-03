#include "stdafx.h"
#include "IESpy.h"
#include "FlashSink2.h"
#include <fstream>
#include <afxctl.h>

using namespace std;

BEGIN_DISPATCH_MAP(CFlashSink2, CCmdTarget)
	DISP_FUNCTION_ID(CFlashSink2, "OnReadyStateChange", DISPID_ONREADYSTATECHANGE, OnReadyStateChange, VT_EMPTY, VTS_I4)
	DISP_FUNCTION_ID(CFlashSink2, "OnProgress", DISPID_ONPROGRESS, OnProgress, VT_EMPTY, VTS_I4)
	DISP_FUNCTION_ID(CFlashSink2, "FSCommand", DISPID_FSCOMMAND, OnFSCommand, VT_EMPTY, VTS_BSTR VTS_BSTR)
	DISP_FUNCTION_ID(CFlashSink2, "FlashCall", DISPID_FLASHCALL, OnFlashCall, VT_EMPTY, VTS_BSTR)
	DISP_FUNCTION_ID(CFlashSink2, "GotFocus", DISPID_GOTFOCUS, OnGotFocus, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION_ID(CFlashSink2, "LostFocus", DISPID_LOSTFOCUS, OnLostFocus, VT_EMPTY, VTS_NONE)
END_DISPATCH_MAP()

void CFlashSink2::OnReadyStateChange(long newState)
{
	REDIR_OUT2LOG
	cout << "CFlashSink2::OnReadyStateChange: " << newState << endl;
	REDIR_OUT2OUT
}

void CFlashSink2::OnProgress(long percentDone)
{
	REDIR_OUT2LOG
	cout << "CFlashSink2::OnProgress: " << percentDone << endl;
	REDIR_OUT2OUT
}

void CFlashSink2::OnFSCommand(BSTR command, BSTR args)
{
	REDIR_OUT2LOG
	cout << "CFlashSink2::OnFSCommand: " << command << _T("(") << args << _T(")") << endl;
	REDIR_OUT2OUT
}

void CFlashSink2::OnFlashCall(BSTR request)
{
	REDIR_OUT2LOG
	cout << "CFlashSink2::OnFlashCall: " << request << endl;
	REDIR_OUT2OUT
}

void CFlashSink2::OnGotFocus()
{
	REDIR_OUT2LOG
	cout << "CFlashSink2::OnGotFocus" << endl;
	REDIR_OUT2OUT
}

void CFlashSink2::OnLostFocus()
{
	REDIR_OUT2LOG
	cout << "CFlashSink2::OnLostFocus" << endl;
	REDIR_OUT2OUT
}