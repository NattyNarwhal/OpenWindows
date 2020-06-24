/*
 * Copyright (c) 2004 Pascal Hurni
 *
 * Permission is hereby granted, free of charge, to any person obtaining
 * a copy of this software and associated documentation files (the
 * "Software"), to deal in the Software without restriction, including
 * without limitation the rights to use, copy, modify, merge, publish,
 * distribute, sublicense, and/or sell copies of the Software, and to
 * permit persons to whom the Software is furnished to do so, subject to
 * the following conditions:
 * 
 * The above copyright notice and this permission notice shall be
 * included in all copies or substantial portions of the Software.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
 * EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
 * MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
 * NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
 * LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
 * OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
 * WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 */

#include "ShellFolderView.h"

// define some undocumented messages. See "shlext.h" from Henk Devos & Andrew Le Bihan, at http://www.whirlingdervishes.com/nselib/public
#define SFVCB_SELECTIONCHANGED    0x0008
#define SFVCB_DRAWMENUITEM        0x0009
#define SFVCB_MEASUREMENUITEM     0x000A
#define SFVCB_EXITMENULOOP        0x000B
#define SFVCB_VIEWRELEASE         0x000C
#define SFVCB_GETNAMELENGTH       0x000D
#define SFVCB_WINDOWCLOSING       0x0010
#define SFVCB_LISTREFRESHED       0x0011
#define SFVCB_WINDOWFOCUSED       0x0012
#define SFVCB_REGISTERCOPYHOOK    0x0014
#define SFVCB_COPYHOOKCALLBACK    0x0015
#define SFVCB_ADDINGOBJECT        0x001D
#define SFVCB_REMOVINGOBJECT      0x001E

#define SFVCB_GETCOMMANDDIR       0x0021
#define SFVCB_GETCOLUMNSTREAM     0x0022
#define SFVCB_CANSELECTALL        0x0023
#define SFVCB_ISSTRICTREFRESH     0x0025
#define SFVCB_ISCHILDOBJECT       0x0026
#define SFVCB_GETEXTVIEWS         0x0028
#define SFVCB_WNDMAIN              46
#define SFVCB_COLUMNCLICK2   0x32


// Macros to trace the message name instead of its id
#define BEGIN_TRACE_MSG_NAME() switch (uMsg) {
#define END_TRACE_MSG_NAME() default: ATLTRACE("COWRootShellView(%08x) Msg: %2d:%2x w=%d, l=%d\n", this, uMsg, uMsg, wParam, lParam); }
#define TRACE_MSG_NAME(name) case name: ATLTRACE("COWRootShellView(%08x) %s\n", this, #name); break;


// This class does very little but it trace the messages.
// You can add message handler like the SFVM_COLUMNCLICK.
class COWRootShellView : public CShellFolderViewImpl
{
public:
	COWRootShellView()
	{
		ATLTRACE("COWRootShellView(%08x) CONSTRUCTOR\n", this);
	}

	~COWRootShellView()
	{
		ATLTRACE("COWRootShellView(%08x) DESTRUCTOR\n", this);
	}

	// If called, the passed object will be held (AddRef()'ed) until the View gets deleted.
	void Init(IUnknown *pUnkOwner = NULL)
	{
		m_UnkOwnerPtr = pUnkOwner;
	}

	// The message map
	BEGIN_MSG_MAP(COWRootShellView)
		BEGIN_TRACE_MSG_NAME()		
			TRACE_MSG_NAME(SFVM_MERGEMENU)
			TRACE_MSG_NAME(SFVM_INVOKECOMMAND)
			TRACE_MSG_NAME(SFVM_GETHELPTEXT)
			TRACE_MSG_NAME(SFVM_GETTOOLTIPTEXT)
			TRACE_MSG_NAME(SFVM_GETBUTTONINFO)
			TRACE_MSG_NAME(SFVM_GETBUTTONS)
			TRACE_MSG_NAME(SFVM_INITMENUPOPUP)
			TRACE_MSG_NAME(SFVM_FSNOTIFY)
			TRACE_MSG_NAME(SFVM_WINDOWCREATED)
			TRACE_MSG_NAME(SFVM_GETDETAILSOF)
			TRACE_MSG_NAME(SFVM_COLUMNCLICK)
			TRACE_MSG_NAME(SFVM_QUERYFSNOTIFY)
			TRACE_MSG_NAME(SFVM_DEFITEMCOUNT)
			TRACE_MSG_NAME(SFVM_DEFVIEWMODE)
			TRACE_MSG_NAME(SFVM_UNMERGEMENU)
			TRACE_MSG_NAME(SFVM_UPDATESTATUSBAR)
			TRACE_MSG_NAME(SFVM_BACKGROUNDENUM)
			TRACE_MSG_NAME(SFVM_DIDDRAGDROP)
			TRACE_MSG_NAME(SFVM_SETISFV)
			TRACE_MSG_NAME(SFVM_THISIDLIST)
			TRACE_MSG_NAME(SFVM_ADDPROPERTYPAGES)
			TRACE_MSG_NAME(SFVM_BACKGROUNDENUMDONE)
			TRACE_MSG_NAME(SFVM_GETNOTIFY)
			TRACE_MSG_NAME(SFVM_GETSORTDEFAULTS)
			TRACE_MSG_NAME(SFVM_SIZE)
			TRACE_MSG_NAME(SFVM_GETZONE)
			TRACE_MSG_NAME(SFVM_GETPANE)
			TRACE_MSG_NAME(SFVM_GETHELPTOPIC)
			TRACE_MSG_NAME(SFVM_GETANIMATION)

			TRACE_MSG_NAME(SFVCB_SELECTIONCHANGED)
			TRACE_MSG_NAME(SFVCB_DRAWMENUITEM)
			TRACE_MSG_NAME(SFVCB_MEASUREMENUITEM)
			TRACE_MSG_NAME(SFVCB_EXITMENULOOP)
			TRACE_MSG_NAME(SFVCB_VIEWRELEASE)
			TRACE_MSG_NAME(SFVCB_GETNAMELENGTH)
			TRACE_MSG_NAME(SFVCB_WINDOWCLOSING)
			TRACE_MSG_NAME(SFVCB_LISTREFRESHED)
			TRACE_MSG_NAME(SFVCB_WINDOWFOCUSED)
			TRACE_MSG_NAME(SFVCB_REGISTERCOPYHOOK)
			TRACE_MSG_NAME(SFVCB_COPYHOOKCALLBACK)
			TRACE_MSG_NAME(SFVCB_ADDINGOBJECT)
			TRACE_MSG_NAME(SFVCB_REMOVINGOBJECT)
			TRACE_MSG_NAME(SFVCB_GETCOMMANDDIR)
			TRACE_MSG_NAME(SFVCB_GETCOLUMNSTREAM)
			TRACE_MSG_NAME(SFVCB_CANSELECTALL)
			TRACE_MSG_NAME(SFVCB_ISSTRICTREFRESH)
			TRACE_MSG_NAME(SFVCB_ISCHILDOBJECT)
			TRACE_MSG_NAME(SFVCB_COLUMNCLICK2)
			TRACE_MSG_NAME(SFVCB_GETEXTVIEWS)
			TRACE_MSG_NAME(SFVCB_WNDMAIN)
		END_TRACE_MSG_NAME()

		MESSAGE_HANDLER(SFVM_COLUMNCLICK, OnColumnClick)
		MESSAGE_HANDLER(SFVM_GETDETAILSOF, OnGetDetailsOf)
		MESSAGE_HANDLER(SFVM_DEFVIEWMODE, OnDefViewMode)
	END_MSG_MAP()

	// Offer to set the default view mode
	LRESULT OnDefViewMode(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL &bHandled)
	{
		ATLTRACE("COWRootShellView(%08x)::OnDefViewMode()\n", this);
#ifdef FVM_CONTENT
		/* Requires Windows 7+, by Gravis' request */
		DWORD ver, maj, min;
		ver = GetVersion();
		maj = (DWORD)(LOBYTE(LOWORD(dwVersion)));
		min = (DWORD)(HIBYTE(LOWORD(dwVersion)));
		if (maj > 6 || (maj == 6 && min >= 1))
			*(FOLDERVIEWMODE*)lParam = FVM_CONTENT;
#endif
		return S_OK;
	}

	// When a user clicks on a column header in details mode
	LRESULT OnColumnClick(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL &bHandled)
	{
		ATLTRACE("COWRootShellView(%08x)::OnColumnClick(iColumn=%d)\n", this, wParam);

		// Shell version 4.7x doesn't understand S_FALSE as described in the SDK.
		SendFolderViewMessage(SFVM_REARRANGE, wParam);
		return S_OK;
	}

	// This message is used with shell version 4.7x, shell 5 and above prefer to use IShellFolder2::GetDetailsOf()
	LRESULT OnGetDetailsOf(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL &bHandled)
	{
		int iColumn = (int)wParam;
		DETAILSINFO* pDi = (DETAILSINFO*)lParam;

		ATLTRACE("COWRootShellView(%08x)::OnGetDetailsOf(iColumn=%d)\n", this, iColumn);

		if (!pDi)
			return E_POINTER;

		HRESULT hr;
		SHELLDETAILS ShellDetails;

		IShellDetails *pISD;
		hr = m_pISF->QueryInterface(IID_IShellDetails, (void**)&pISD);
		if (FAILED(hr))
			return hr;

		hr = pISD->GetDetailsOf(pDi->pidl, iColumn, &ShellDetails);
		pISD->Release();
		if (FAILED(hr))
			return hr;

		pDi->cxChar = ShellDetails.cxChar;
		pDi->fmt = ShellDetails.fmt;
		pDi->str = ShellDetails.str;
		pDi->iImage = 0;

		return S_OK;
	}

protected:
	CComPtr<IUnknown> m_UnkOwnerPtr;
};

