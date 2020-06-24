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

//========================================================================================
// Implements a IShellView like requested by IShellFolder::CreateViewObject().
// Derive from this class and add an ATL message map to handle the SFVM_* messages.
// When request to IShellView is made, construct your derived object and call Create()
// with the first param to the storage receiving the new IShellView.
//
// Example:
//	CRootShellView is a CShellFolderViewImpl subclass
/*
	STDMETHODIMP CRootShellFolder::CreateViewObject(HWND hwndOwner, REFIID riid, void** ppvOut)
	{
		HRESULT hr;

		if (ppvOut == NULL)
			return E_POINTER;
		*ppvOut = NULL;

		if (riid == IID_IShellView)
		{
			// Create a view object
			CComObject<CRootShellView>* pViewObject;
			hr = CComObject<CRootShellView>::CreateInstance(&pViewObject);
			if (FAILED(hr))
				return hr;

			// AddRef the object while we are using it
			pViewObject->AddRef();

			// Create the view
			hr = pViewObject->Create((IShellView**)ppvOut, hwndOwner, (IShellFolder*)this);

			// We are finished with our own use of the view object (AddRef()'ed above by us, AddRef()'ed by Create)
			pViewObject->Release();

			return hr;
		}

		return E_NOINTERFACE;
	}
*/
// ATL message maps:
//	If you did not handled the message set bHandled to FALSE. (Defaults to TRUE when your message handler is called)
//	This will return E_NOTIMPL to the caller which means that the message was not handled.
//	When you have handled the message, return (as LRESULT) S_OK (which is 0) or any other valid value described in
//	the SDK for your message.
//
//	In any of your message handler, you can use m_pISF which is a pointer to the related IShellFolder.
//
//========================================================================================


#ifndef __SHELLFOLDERVIEW_H__
#define __SHELLFOLDERVIEW_H__

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000


//========================================================================================

class CShellFolderViewImpl : public CMessageMap, public CComObjectRoot, public IShellFolderViewCB
{
public:

	// This really creates the view.
	// ppISV		will receive the interface pointer to the view (that one should be returned from IShellFolder::CreateViewObject)
	// hwndOwner	The window handle of the parent to the new shell view
	// pISF			The IShellFolder related to the view
	HRESULT Create(IShellView **ppISV, HWND hwndOwner, IShellFolder *pISF, IShellView *psvOuter = NULL)
	{
		m_hwndOwner = hwndOwner;

		SFV_CREATE sfv;
		sfv.cbSize = sizeof(sfv);
		sfv.pshf = pISF;
		sfv.psvOuter = psvOuter;
		sfv.psfvcb = (IShellFolderViewCB*)this;

		m_pISF = pISF;

		return SHCreateShellFolderView(&sfv, ppISV);
	}

	// Used to send messages back to the shell view
	LRESULT SendFolderViewMessage(UINT uMsg, LPARAM lParam)
	{
		return SHShellFolderView_Message(m_hwndOwner, uMsg, lParam);
	}

public:
	// Implementation

	BEGIN_COM_MAP(CShellFolderViewImpl)
		COM_INTERFACE_ENTRY_IID(IID_IShellFolderViewCB, IShellFolderViewCB)
	END_COM_MAP()

	STDMETHODIMP MessageSFVCB(UINT uMsg, WPARAM wParam, LPARAM lParam)
	{
		LRESULT lResult;
		BOOL bResult = ProcessWindowMessage(NULL, uMsg, wParam, lParam, lResult, 0);
		return bResult ? lResult : E_NOTIMPL;
	}

protected:
	HWND m_hwndOwner;
	IShellFolder *m_pISF;		// This one is not ref-counted. This object should be garanted to live
								// until the view is destroyed. So the lifetime is handled by SHCreateShellFolderView()
};

//========================================================================================

#endif // __SHELLFOLDERVIEW_H__

