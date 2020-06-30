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

#ifndef __SHELLITEMS_H_
#define __SHELLITEMS_H_

#include "MPidlMgr.h"
#include "CStringCopyTo.h"
using namespace Mortimer;


//========================================================================================

// Used by SetReturnString.
// Can also be used by anyone when the Shell Allocator is needed. Use the gloabl g_Malloc object which is a CMalloc.
struct CMalloc
{
	CComPtr<IMalloc> m_MallocPtr;
	CMalloc();
};

// Set the return string 'Source' in the STRRET struct.
// Note that it always allocate a UNICODE copy of the string.
// Returns false if memory allocation fails.
bool SetReturnStringA(LPCSTR Source, STRRET &str);
bool SetReturnStringW(LPCWSTR Source, STRRET &str);

#ifdef _UNICODE
	#define SetReturnString SetReturnStringW
#else
	#define SetReturnString SetReturnStringA
#endif


//========================================================================================
// This class handles our data that gets embedded in a pidl.

class COWItem : public CPidlData
{
public:

	//-------------------------------------------------------------------------------
	// used by the manager to embed data, previously set by clients, into a pidl

	// The pidl signature
	enum { MAGIC = 0xAA000055 | ('OW'<<8) };

	// return the size of the pidl data. Not counting the mkid.cb member.
	ULONG GetSize();

	// copy the data to the target
	void CopyTo(void *pTarget);

	//-------------------------------------------------------------------------------
	// Used by clients to set data

	// The target path
	void SetPath(LPCWSTR Path);

	// The display name (may contain any chars)
	void SetName(LPCWSTR Name);

	// The rank (preferred items get low numbers, starting at 1)
	void SetRank(USHORT Rank);

	//-------------------------------------------------------------------------------
	// Used by clients to get data from a given pidl

	// Is this pidl really one of ours?
	static bool IsOwn(LPCITEMIDLIST pidl);

	// Retrieve target path.
	// The pidl MUST remain valid until the caller has finished with the returned string.
	static LPOLESTR GetPath(LPCITEMIDLIST pidl);

	// Retrieve display name.
	// The pidl MUST remain valid until the caller has finished with the returned string.
	static LPOLESTR GetName(LPCITEMIDLIST pidl);

	// Retrieve the item rank
	static USHORT GetRank(LPCITEMIDLIST pidl);

	//-------------------------------------------------------------------------------

protected:
	USHORT m_Padding; /* Pascal's example used an item type here, we don't care about that */
	USHORT m_Rank;
	// The old wtlstr CString is always TCHAR, not a templated version, and
	// do we really want to reimplement that? For now, statically allocate
	// MAX_PATH worth. On Windows 10, this can be larger, but we can truncate
	// for now.
	wchar_t m_Path[MAX_PATH];
	wchar_t m_Name[MAX_PATH];
};


// Collection for our data
typedef CSimpleArray<COWItem> COWItemList;

//========================================================================================
// Light implementation of IDataObject.
//
// This object is used when you double-click on an item in the FileDialog.
// It's purpose is simply to encapsulate the complete pidl for the item (remember it's a Favorite item)
// into the IDataObject, so that the FileDialog can pass it further to our IShellFolder::BindToObject().
// Because I'm only interested in the FileDialog behaviour, every methods returns E_NOTIMPL except GetData().

class ATL_NO_VTABLE CDataObject :
	public CComObjectRootEx<CComSingleThreadModel>,
	public IDataObject, public IEnumFORMATETC
{
public:
	BEGIN_COM_MAP(CDataObject)
		COM_INTERFACE_ENTRY_IID(IID_IDataObject, IDataObject)
		COM_INTERFACE_ENTRY_IID(IID_IEnumFORMATETC, IEnumFORMATETC)
	END_COM_MAP()

	//-------------------------------------------------------------------------------

	CDataObject();
	~CDataObject();

	// Ensure the owner object is not freed before this one
	void Init(IUnknown *pUnkOwner);

	// Populate the object with the Favorite Item pidl.
	// This member must be called before any IDataObject member.
	void SetPidl(LPCITEMIDLIST pidlParent, LPCITEMIDLIST pidl);

	//-------------------------------------------------------------------------------
	// IDataObject methods

	STDMETHOD(GetData) (LPFORMATETC pFE, LPSTGMEDIUM pStgMedium);
	STDMETHOD(GetDataHere) (LPFORMATETC, LPSTGMEDIUM);
	STDMETHOD(QueryGetData) (LPFORMATETC);
	STDMETHOD(GetCanonicalFormatEtc) (LPFORMATETC, LPFORMATETC);
	STDMETHOD(SetData) (LPFORMATETC, LPSTGMEDIUM, BOOL);
	STDMETHOD(EnumFormatEtc) (DWORD, IEnumFORMATETC**);
	STDMETHOD(DAdvise) (LPFORMATETC, DWORD, IAdviseSink*, LPDWORD);
	STDMETHOD(DUnadvise) (DWORD dwConnection);
	STDMETHOD(EnumDAdvise) (IEnumSTATDATA** ppEnumAdvise);

	//-------------------------------------------------------------------------------
	// IEnumFORMATETC members

	STDMETHOD(Next) (ULONG, LPFORMATETC, ULONG*);
	STDMETHOD(Skip) (ULONG);
	STDMETHOD(Reset) ();
	STDMETHOD(Clone) (LPENUMFORMATETC*);

protected:
	CComPtr<IUnknown> m_UnkOwnerPtr;
	CPidlMgr m_PidlMgr;

	UINT m_cfShellIDList;

	LPITEMIDLIST m_pidl;
	LPITEMIDLIST m_pidlParent;
};


#endif // __SHELLITEMS_H_
