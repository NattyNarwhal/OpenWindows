/*
 * Copyright (c) 2004 Pascal Hurni
 * Copyright (c) 2020 Calvin Buckley
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

#include "stdafx.h"
#include "OpenWindows_h.h"
#include "RootShellFolder.h"

#include "RootShellView.h"


//========================================================================================
// Helpers

#ifdef _DEBUG

LPTSTR PidlToString(LPCITEMIDLIST pidl)
{
	static int which = 0;
	static TCHAR str1[200];
	static TCHAR str2[200];
	TCHAR *str;
	which ^= 1;
	str = which ? str1 : str2;

	str[0] = '\0';

	if (pidl == NULL)
	{
		_tcscpy(str, _T("<null>"));
		return str;
	}

	while (pidl->mkid.cb != 0)
	{
		if (*((DWORD*)(pidl->mkid.abID)) == COWItem::MAGIC)
		{
#ifndef _UNICODE
			char tmp[128];
			mbstowcs((USHORT*)(pidl->mkid.abID)+4, tmp, 128);
			_tcscat(str, tmp);
#else
			_tcscat(str, (wchar_t*)(pidl->mkid.abID)+4);
#endif
			_tcscat(str, _T("::"));
		}
		else
		{
			TCHAR tmp[16];
			_stprintf(tmp, _T("<unk-%02d>::"), pidl->mkid.cb);
			_tcscat(str, tmp);
		}

		pidl = LPITEMIDLIST(LPBYTE(pidl) + pidl->mkid.cb);
	}
	return str;
}

inline LPOLESTR iidToString(REFIID iid)
{
	LPOLESTR str;
	StringFromIID(iid, &str);
	return str;						// TODO: BUG: This string SHOULD be freed, but it isn't !
}

#define DUMPIID(iid) AtlDumpIID(iid, NULL, S_OK)

#else
#define PidlToString
#define iidToString
#define DUMPIID
#endif

//========================================================================================
// Helper class for the CComEnumOnCArray

class CCopyItemPidl
{
public:
	static void init(LPITEMIDLIST* p)
	{
	}

	static HRESULT copy(LPITEMIDLIST* pTo, COWItem* pFrom)
	{
		*pTo = s_PidlMgr.Create(*pFrom);
		return (NULL != *pTo) ? S_OK : E_OUTOFMEMORY;
	}

	static void destroy(LPITEMIDLIST* p)
	{
		s_PidlMgr.Delete(*p); 
	}

protected:
	static CPidlMgr s_PidlMgr;
};

CPidlMgr CCopyItemPidl::s_PidlMgr;

// This class implements the IEnumIDList for our CDataFavo items.
typedef CComEnumOnCArray<IEnumIDList, &IID_IEnumIDList, LPITEMIDLIST, CCopyItemPidl, COWItemList> CEnumItemsIDList;

//========================================================================================
// COWRootShellFolder

COWRootShellFolder::COWRootShellFolder() : m_pidlRoot(NULL)
{

}

STDMETHODIMP COWRootShellFolder::GetClassID(CLSID* pClsid)
{
	if ( NULL == pClsid )
		return E_POINTER;

	// Return our GUID to the shell.
	*pClsid = CLSID_OpenWindowsRootShellFolder;

	return S_OK;
}

// Initialize() is passed the PIDL of the folder where our extension is.
STDMETHODIMP COWRootShellFolder::Initialize(LPCITEMIDLIST pidl)
{
	ATLTRACE(_T("COWRootShellFolder(0x%08x)::Initialize() pidl=[%s]\n"), this, PidlToString(pidl));

	m_pidlRoot = m_PidlMgr.Copy(pidl);

	return S_OK;
}

STDMETHODIMP COWRootShellFolder::GetCurFolder(LPITEMIDLIST *ppidl)
{
	ATLTRACE(_T("COWRootShellFolder(0x%08x)::GetCurFolder()\n"), this);

	if (ppidl == NULL)
		return E_POINTER;

	*ppidl = m_PidlMgr.Copy(m_pidlRoot);

	return S_OK;
}

//-------------------------------------------------------------------------------
// IShellFolder

// BindToObject() is called when a folder in our part of the namespace is being browsed.
STDMETHODIMP COWRootShellFolder::BindToObject(LPCITEMIDLIST pidl, LPBC pbcReserved, REFIID riid, void** ppvOut)
{
	ATLTRACE(_T("COWRootShellFolder(0x%08x)::BindToObject() pidl=[%s]\n"), this, PidlToString(pidl));

	// If the passed pidl is not ours, fail.
	if (!COWItem::IsOwn(pidl))
		return E_INVALIDARG;

	CComPtr<IShellFolder> DesktopPtr;

	// We only support single-level pidl (coz subitems are, in reality, plain path to other locations)
	// BUT, the modified Office FileDialog still uses our root IShellFolder to request bindings for sub-sub-items, so we do it!
	// Note that we also use multi-level pidl if "EnableFavoritesSubfolders" is enabled.
	if (!m_PidlMgr.IsSingle(pidl))
	{
		HRESULT hr;
		hr = SHGetDesktopFolder(&DesktopPtr);
		if (FAILED(hr))
			return hr;

		LPITEMIDLIST pidlLocal;
		hr = DesktopPtr->ParseDisplayName(NULL, pbcReserved, COWItem::GetPath(pidl), NULL, &pidlLocal, NULL);
		if (FAILED(hr))
			return hr;

		// Bind to the root folder of the favorite folder
		CComPtr<IShellFolder> RootFolderPtr;
		hr = DesktopPtr->BindToObject(pidlLocal, NULL, IID_IShellFolder, (void**)&RootFolderPtr);
		ILFree(pidlLocal);
		if (FAILED(hr))
			return hr;

		// And now bind to the sub-item of it
		return RootFolderPtr->BindToObject(m_PidlMgr.GetNextItem(pidl), pbcReserved, riid, ppvOut);
	}

	// Okay, browsing into a favorite item will redirect to its real path.
	HRESULT hr;
	hr = SHGetDesktopFolder(&DesktopPtr);
	if (FAILED(hr))
		return hr;

	LPITEMIDLIST pidlLocal;
	hr = DesktopPtr->ParseDisplayName(NULL, pbcReserved, COWItem::GetPath(pidl), NULL, &pidlLocal, NULL);
	if (FAILED(hr))
		return hr;

	hr = DesktopPtr->BindToObject(pidlLocal, pbcReserved, riid, ppvOut);

	ILFree(pidlLocal);
	return hr;

	// could also use this one? ILCreateFromPathW 
}

// CompareIDs() is responsible for returning the sort order of two PIDLs.
// lParam can be the 0-based Index of the details column
STDMETHODIMP COWRootShellFolder::CompareIDs(LPARAM lParam, LPCITEMIDLIST pidl1, LPCITEMIDLIST pidl2)
{
	ATLTRACE(_T("COWRootShellFolder(0x%08x)::CompareIDs(lParam=%d) pidl1=[%s], pidl2=[%s]\n"), this, lParam, PidlToString(pidl1), PidlToString(pidl2));

	// First check if the pidl are ours
	if (!COWItem::IsOwn(pidl1) || !COWItem::IsOwn(pidl2))
		return E_INVALIDARG;

	// Now check if the pidl are one or multi level, in case they are multi-level, return non-equality
	if (!m_PidlMgr.IsSingle(pidl1) || !m_PidlMgr.IsSingle(pidl2))
		return MAKE_HRESULT(SEVERITY_SUCCESS, 0, 1);

	USHORT Result = 0;	// see note below (MAKE_HRESULT)

	switch (lParam & SHCIDS_COLUMNMASK)
	{
	case DETAILS_COLUMN_NAME:		Result = wcscmp(COWItem::GetName(pidl1), COWItem::GetName(pidl2));	break;
	case DETAILS_COLUMN_PATH:		Result = wcscmp(COWItem::GetPath(pidl1), COWItem::GetPath(pidl2));	break;
	case DETAILS_COLUMN_RANK:		Result = COWItem::GetRank(pidl1)-COWItem::GetRank(pidl2);			break;
	default:						return E_INVALIDARG;
	}

	// Warning: the last param MUST be unsigned, if not (ie: short) a negative value will trash the high order word of the HRESULT!
	return MAKE_HRESULT(SEVERITY_SUCCESS, 0, /*-1,0,1*/Result);
}

// CreateViewObject() creates a new COM object that implements IShellView.
STDMETHODIMP COWRootShellFolder::CreateViewObject(HWND hwndOwner, REFIID riid, void** ppvOut)
{
	ATLTRACE(_T("COWRootShellFolder(0x%08x)::CreateViewObject()"), this);
	//DUMPIID(riid);

	HRESULT hr;

	if (ppvOut == NULL)
		return E_POINTER;

	*ppvOut = NULL;

	// We handle only the IShellView
	if (riid == IID_IShellView)
	{
		ATLTRACE(_T(" ** CreateViewObject for IShellView"));

		// Create a view object
		CComObject<COWRootShellView>* pViewObject;
		hr = CComObject<COWRootShellView>::CreateInstance(&pViewObject);
		if (FAILED(hr))
			return hr;

		// AddRef the object while we are using it
		pViewObject->AddRef();

		// Tight the view object lifetime with the current IShellFolder.
		pViewObject->Init(GetUnknown());

		// Create the view
		hr = pViewObject->Create((IShellView**)ppvOut, hwndOwner, (IShellFolder*)this);

		// We are finished with our own use of the view object (AddRef()'ed above by us, AddRef()'ed by Create)
		pViewObject->Release();

		return hr;
	}
	
#if _DEBUG
	// vista sludge that no one knows what it actually is
	static const GUID unknownVistaGuid = // {93F81976-6A0D-42C3-94DD-AA258A155470}
    {0x93F81976, 0x6A0D, 0x42C3, {0x94, 0xDD, 0xAA, 0x25, 0x8A, 0x15, 0x54, 0x70}};

	if (riid != unknownVistaGuid) {
		LPOLESTR unkIid = iidToString(riid);
		ATLTRACE(_T(" ** CreateViewObject is unknown: %s"), unkIid);
		CoTaskMemFree(unkIid);
	}
#endif

	// We do not handle other objects
	return E_NOINTERFACE;
}

// EnumObjects() creates a COM object that implements IEnumIDList.
STDMETHODIMP COWRootShellFolder::EnumObjects(HWND hwndOwner, DWORD dwFlags, LPENUMIDLIST* ppEnumIDList)
{
	ATLTRACE("COWRootShellFolder(0x%08x)::EnumObjects(dwFlags=0x%04x)\n", this, dwFlags);

	HRESULT hr;

	if (ppEnumIDList == NULL)
		return E_POINTER;

    *ppEnumIDList = NULL;

    // Enumerate DOpus Favorites and put them in an array
	m_OpenedWindows.RemoveAll();

	EnumerateExplorerWindows(&m_OpenedWindows, hwndOwner);

	ATLTRACE(_T(" ** EnumObjects: Now have %d items"), m_OpenedWindows.GetSize());

    // Create an enumerator with CComEnumOnCArray<> and our copy policy class.
	CComObject<CEnumItemsIDList>* pEnum;
	hr = CComObject<CEnumItemsIDList>::CreateInstance(&pEnum);
	if (FAILED(hr))
		return hr;

    // AddRef() the object while we're using it.
	pEnum->AddRef();

    // Init the enumerator.  Init() will AddRef() our IUnknown (obtained with
    // GetUnknown()) so this object will stay alive as long as the enumerator 
    // needs access to the collection m_Favorites.
	hr = pEnum->Init(GetUnknown(), m_OpenedWindows);

    // Return an IEnumIDList interface to the caller.
	if (SUCCEEDED(hr))
		hr = pEnum->QueryInterface(IID_IEnumIDList, (void**)ppEnumIDList);

	pEnum->Release();

	return hr;
}


// GetAttributesOf() returns the attributes for the items whose PIDLs are passed in.
STDMETHODIMP COWRootShellFolder::GetAttributesOf(UINT uCount, LPCITEMIDLIST aPidls[], LPDWORD pdwAttribs)
{
#ifdef _DEBUG
	if (uCount >= 1)
		ATLTRACE(_T("COWRootShellFolder(0x%08x)::GetAttributesOf(uCount=%d) pidl=[%s]\n"), this, uCount, PidlToString(aPidls[0]));
	else
		ATLTRACE("COWRootShellFolder(0x%08x)::GetAttributesOf(uCount=%d)\n", this, uCount);
#endif

	// We limit the tree, by indicating that the favorites folder does not contain sub-folders

	if ((uCount == 0) || (aPidls[0]->mkid.cb == 0))
	    *pdwAttribs &= SFGAO_HASSUBFOLDER|SFGAO_FOLDER | SFGAO_FILESYSTEM|SFGAO_FILESYSANCESTOR | SFGAO_BROWSABLE;
	else 
	    *pdwAttribs &= SFGAO_FOLDER | SFGAO_FILESYSTEM|SFGAO_FILESYSANCESTOR | SFGAO_BROWSABLE | SFGAO_LINK;

    return S_OK;
}


// GetUIObjectOf() is called to get several sub-objects like IExtractIcon and IDataObject
STDMETHODIMP COWRootShellFolder::GetUIObjectOf(HWND hwndOwner, UINT uCount, LPCITEMIDLIST* pPidl, REFIID riid, LPUINT puReserved, void** ppvReturn)
{
#ifdef _DEBUG
	if (uCount >= 1)
		ATLTRACE(_T("COWRootShellFolder(0x%08x)::GetUIObjectOf(uCount=%d) pidl=[%s]"), this, uCount, PidlToString(*pPidl));
	else
		ATLTRACE(_T("COWRootShellFolder(0x%08x)::GetUIObjectOf(uCount=%d)"), this, uCount);
	//DUMPIID(riid);
#endif

	HRESULT hr;

	if (ppvReturn == NULL)
		return E_POINTER;

	*ppvReturn = NULL;

	if (uCount == 0)
		return E_INVALIDARG;

	// Does the FileDialog need to embed some data?
	if (riid == IID_IDataObject)
	{
		// Only one item at a time
		if (uCount != 1)
			return E_INVALIDARG;

		// Is this really one of our item?
		if (!COWItem::IsOwn(*pPidl))
			return E_INVALIDARG;

		// Create a COM object that exposes IDataObject
		CComObject<CDataObject>* pDataObject;
		hr = CComObject<CDataObject>::CreateInstance(&pDataObject);
		if (FAILED(hr))
			return hr;

		// AddRef it while we are working with it, this prevent from an early destruction.
		pDataObject->AddRef();

		// Tight its lifetime with this object (the IShellFolder object)
		pDataObject->Init(GetUnknown());

		// Okay, embed the pidl in the data
		pDataObject->SetPidl(m_pidlRoot, *pPidl);

		// Return the requested interface to the caller
        hr = pDataObject->QueryInterface(riid, ppvReturn);

		// We do no more need our ref (note that the object will not die because the QueryInterface above, AddRef'd it)
		pDataObject->Release();
		return hr;
	}

	// All other requests are delegated to the target path's IShellFolder

	// because multiple items can point to different storages, we can't (easily) handle groups of items.
	if (uCount > 1)
		return E_NOINTERFACE;

	CComPtr<IShellFolder> TargetParentShellFolderPtr;
	CComPtr<IShellFolder> DesktopPtr;

	hr = SHGetDesktopFolder(&DesktopPtr);
	if (FAILED(hr))
		return hr;

	LPITEMIDLIST pidlLocal;
	hr = DesktopPtr->ParseDisplayName(NULL, NULL, COWItem::GetPath(*pPidl), NULL, &pidlLocal, NULL);
	if (FAILED(hr))
		return hr;

	LPITEMIDLIST pidlRelative;

	//------------------------------
	// this block emulate the following line (not available to shell version 4.7x)
	//		hr = SHBindToParent(pidlLocal, IID_IShellFolder, (void**)&pTargetParentShellFolder, &pidlRelative);
	LPITEMIDLIST pidlTmp = ILFindLastID(pidlLocal);
	pidlRelative = ILClone(pidlTmp);
	ILRemoveLastID(pidlLocal);
	hr = DesktopPtr->BindToObject(pidlLocal, NULL, IID_IShellFolder, (void**)&TargetParentShellFolderPtr);
	ILFree(pidlLocal);
	if (FAILED(hr))
	{
		ILFree(pidlRelative);
		return hr;
	}
	//------------------------------

	hr = TargetParentShellFolderPtr->GetUIObjectOf(hwndOwner, 1, (LPCITEMIDLIST*)&pidlRelative, riid, puReserved, ppvReturn);

	ILFree(pidlRelative);

	return hr;
}

STDMETHODIMP COWRootShellFolder::BindToStorage(LPCITEMIDLIST, LPBC, REFIID, void**)
{
	ATLTRACE("COWRootShellFolder(0x%08x)::BindToStorage()\n", this);
	return E_NOTIMPL;
}

STDMETHODIMP COWRootShellFolder::GetDisplayNameOf(LPCITEMIDLIST pidl, DWORD uFlags, LPSTRRET lpName)
{
	ATLTRACE(_T("COWRootShellFolder(0x%08x)::GetDisplayNameOf(uFlags=0x%04x) pidl=[%s]\n"), this, uFlags, PidlToString(pidl));

	if ((pidl == NULL) || (lpName == NULL))
		return E_POINTER;

	// Return name of Root
	if (pidl->mkid.cb == 0)
	{
		switch (uFlags)
		{
		case SHGDN_NORMAL | SHGDN_FORPARSING :	// <- if wantsFORPARSING is present in the regitry
			// As stated in the SDK, we should return here our virtual junction point which is in the form "::{GUID}"
			// So we should return "::{E477F21A-D9F6-4B44-AD43-A95D622D2910}". This works (well I guess it's never used) great
			// except for the modified FileDialog of Office (97, 2000). Thanx to MS, they always have to tweak their own app.
			// The drawback is that Office WILL check for real filesystem existance of the returned string. As you know, "::{GUID}"
			// is not a real filesystem path. This stops Office FileDialog from browsing our namespace extension!
			// To workaround it, instead returning "::{GUID}", we return a real filesystem path, which will never be used by us
			// nor by the FileDialog. I choosed to return the system temporary directory, which ought te be valid on every system.
			TCHAR TempPath[MAX_PATH];
			if (GetTempPath(MAX_PATH, TempPath) == 0)
				return E_FAIL;

			return SetReturnString(TempPath, *lpName) ? S_OK : E_FAIL;

			// See note above
			// return SetReturnString(_T("::{E477F21A-D9F6-4B44-AD43-A95D622D2910}"), *lpName) ? S_OK : E_FAIL;
		}
		// We dont' handle other combinations of flags
		return E_FAIL;
	}

	// At this stage, the pidl should be one of ours
	if (!COWItem::IsOwn(pidl))
		return E_INVALIDARG;

	switch (uFlags)
	{
	case SHGDN_NORMAL | SHGDN_FORPARSING :
	case SHGDN_INFOLDER | SHGDN_FORPARSING :
		return SetReturnStringW(COWItem::GetPath(pidl), *lpName) ? S_OK : E_FAIL;

	case SHGDN_NORMAL | SHGDN_FOREDITING :
	case SHGDN_INFOLDER | SHGDN_FOREDITING :
		return E_FAIL;	// Can't rename!
	}

	// Any other combination results in returning the name.
	return SetReturnStringW(COWItem::GetName(pidl), *lpName) ? S_OK : E_FAIL;
}

STDMETHODIMP COWRootShellFolder::ParseDisplayName(HWND, LPBC, LPOLESTR, LPDWORD, LPITEMIDLIST*, LPDWORD)
{
	ATLTRACE("COWRootShellFolder(0x%08x)::ParseDisplayName()\n", this);
	return E_NOTIMPL;
}

STDMETHODIMP COWRootShellFolder::SetNameOf(HWND, LPCITEMIDLIST, LPCOLESTR, DWORD, LPITEMIDLIST*)
{
	ATLTRACE("COWRootShellFolder(0x%08x)::SetNameOf()\n", this);
	return E_NOTIMPL;
}

//-------------------------------------------------------------------------------
// IShellDetails

STDMETHODIMP COWRootShellFolder::ColumnClick(UINT iColumn)
{
	ATLTRACE("COWRootShellFolder(0x%08x)::ColumnClick(iColumn=%d)\n", this, iColumn);

	// The caller must sort the column itself
	return S_FALSE;
}


STDMETHODIMP COWRootShellFolder::GetDetailsOf(LPCITEMIDLIST pidl, UINT iColumn, LPSHELLDETAILS pDetails)
{
	ATLTRACE(_T("COWRootShellFolder(0x%08x)::GetDetailsOf(iColumn=%d) pidl=[%s]\n"), this, iColumn, PidlToString(pidl));

	if (iColumn >= DETAILS_COLUMN_MAX)
		return E_FAIL;

	// Shell asks for the column headers
	if (pidl == NULL)
	{
		// Load the iColumn based string from the resource
		CString ColumnName(MAKEINTRESOURCE(IDS_COLUMN_NAME+iColumn));

		if (iColumn == DETAILS_COLUMN_RANK)
		{
			pDetails->fmt = LVCFMT_RIGHT;
			pDetails->cxChar = 6;
		}
		else
		{
			pDetails->fmt = LVCFMT_LEFT;
			pDetails->cxChar = 32;
		}
		return SetReturnString(ColumnName, pDetails->str) ? S_OK : E_OUTOFMEMORY;
	}

	// Okay, this time it's for a real item
	TCHAR tmpStr[16];
	switch (iColumn)
	{
	case DETAILS_COLUMN_NAME:
		pDetails->fmt = LVCFMT_LEFT;
		pDetails->cxChar = wcslen(COWItem::GetName(pidl));
		return SetReturnStringW(COWItem::GetName(pidl), pDetails->str) ? S_OK : E_OUTOFMEMORY;

	case DETAILS_COLUMN_PATH:
		pDetails->fmt = LVCFMT_LEFT;
		pDetails->cxChar = wcslen(COWItem::GetName(pidl));
		return SetReturnStringW(COWItem::GetPath(pidl), pDetails->str) ? S_OK : E_OUTOFMEMORY;
	
	case DETAILS_COLUMN_RANK:
		pDetails->fmt = LVCFMT_RIGHT;
		pDetails->cxChar = 6;
		wsprintf(tmpStr, _T("%d"), COWItem::GetRank(pidl));
		return SetReturnString(tmpStr, pDetails->str) ? S_OK : E_OUTOFMEMORY;
	}

	return E_INVALIDARG;
}

//-------------------------------------------------------------------------------
// IShellFolder2

STDMETHODIMP COWRootShellFolder::EnumSearches(IEnumExtraSearch **ppEnum)
{
	ATLTRACE("COWRootShellFolder(0x%08x)::EnumSearches()\n", this);
	return E_NOTIMPL;
}

STDMETHODIMP COWRootShellFolder::GetDefaultColumn(DWORD dwReserved, ULONG *pSort, ULONG *pDisplay)
{
	ATLTRACE("COWRootShellFolder(0x%08x)::GetDefaultColumn()\n", this);

	if (!pSort || !pDisplay)
		return E_POINTER;

	*pSort = DETAILS_COLUMN_RANK;
	*pDisplay = DETAILS_COLUMN_NAME;

	return S_OK;
}

STDMETHODIMP COWRootShellFolder::GetDefaultColumnState(UINT iColumn, SHCOLSTATEF *pcsFlags)
{
	ATLTRACE("COWRootShellFolder(0x%08x)::GetDefaultColumnState(iColumn=%d)\n", this, iColumn);

	if (!pcsFlags)
		return E_POINTER;

	// Seems that SHCOLSTATE_PREFER_VARCMP doesn't have any noticeable effect (if supplied or not) for Win2K, but don't
	// set it for WinXP, since it will not sort the column. (not setting it means that our CompareIDs() will be called)
	switch (iColumn)
	{
	case DETAILS_COLUMN_NAME:		*pcsFlags = SHCOLSTATE_TYPE_STR | SHCOLSTATE_ONBYDEFAULT;	break;
	case DETAILS_COLUMN_PATH:		*pcsFlags = SHCOLSTATE_TYPE_STR | SHCOLSTATE_ONBYDEFAULT;	break;
	case DETAILS_COLUMN_RANK:		*pcsFlags = SHCOLSTATE_TYPE_INT | SHCOLSTATE_ONBYDEFAULT;	break;
	default:						return E_INVALIDARG;
	}

	return S_OK;
}

STDMETHODIMP COWRootShellFolder::GetDefaultSearchGUID(GUID *pguid)
{
	ATLTRACE("COWRootShellFolder(0x%08x)::GetDefaultSearchGUID()\n", this);
	return E_NOTIMPL;
}

STDMETHODIMP COWRootShellFolder::GetDetailsEx(LPCITEMIDLIST pidl, const SHCOLUMNID *pscid, VARIANT *pv)
{
	ATLTRACE(_T("COWRootShellFolder(0x%08x)::GetDetailsEx(pscid->pid=%d) pidl=[%s]\n"), this, pscid->pid, PidlToString(pidl));

#if defined(OW_PKEYS_SUPPORT)
	PROPVARIANT var;
	/* Vista required */
	if (IsEqualPropertyKey(*pscid, PKEY_PropList_TileInfo))
    {
		ATLTRACE(_T(" ** GetDetailsEx: PKEY_PropList_TileInfo"));
		if (SUCCEEDED(InitPropVariantFromString(_T("prop:System.ItemPathDisplay"), &var)))
		{
			PropVariantToVariant(&var, pv);
			return S_OK;
		}
    }
	else if (IsEqualPropertyKey(*pscid, PKEY_PropList_ExtendedTileInfo))
    {
		ATLTRACE(_T(" ** GetDetailsEx: PKEY_PropList_ExtendedTileInfo"));
		if (SUCCEEDED(InitPropVariantFromString(_T("prop:System.ItemPathDisplay"), &var)))
		{
			PropVariantToVariant(&var, pv);
			return S_OK;
		}
    }
	else if (IsEqualPropertyKey(*pscid, PKEY_PropList_PreviewDetails))
    {
		ATLTRACE(_T(" ** GetDetailsEx: PKEY_PropList_PreviewDetails"));
		if (SUCCEEDED(InitPropVariantFromString(_T("prop:System.ItemPathDisplay"), &var)))
		{
			PropVariantToVariant(&var, pv);
			return S_OK;
		}
    }
	else if (IsEqualPropertyKey(*pscid, PKEY_PropList_FullDetails))
    {
		ATLTRACE(_T(" ** GetDetailsEx: PKEY_PropList_FullDetails"));
		if (SUCCEEDED(InitPropVariantFromString(_T("prop:System.ItemNameDisplay;System.ItemPathDisplay"), &var)))
		{
			PropVariantToVariant(&var, pv);
			return S_OK;
		}
    }
	else if (IsEqualPropertyKey(*pscid, PKEY_ItemType))
    {
		ATLTRACE(_T(" ** GetDetailsEx: PKEY_ItemType"));
		if (SUCCEEDED(InitPropVariantFromString(_T("Directory"), &var)))
		{
			PropVariantToVariant(&var, pv);
			return S_OK;
		}
    }
	// XXX: Are these necessary?
	else if (IsEqualPropertyKey(*pscid, PKEY_ItemNameDisplay))
    {
		ATLTRACE(_T(" ** GetDetailsEx: PKEY_ItemNameDisplay"));
		if (SUCCEEDED(InitPropVariantFromString(COWItem::GetName(pidl), &var)))
		{
			PropVariantToVariant(&var, pv);
			return S_OK;
		}
    }
	else if (IsEqualPropertyKey(*pscid, PKEY_ItemPathDisplay))
    {
		ATLTRACE(_T(" ** GetDetailsEx: PKEY_ItemPathDisplay"));
		if (SUCCEEDED(InitPropVariantFromString(COWItem::GetPath(pidl), &var)))
		{
			PropVariantToVariant(&var, pv);
			return S_OK;
		}
    }
#endif

	ATLTRACE(_T(" ** GetDetailsEx: Not implemented"));
	return E_NOTIMPL;
}

STDMETHODIMP COWRootShellFolder::MapColumnToSCID(UINT iColumn, SHCOLUMNID *pscid)
{
	ATLTRACE("COWRootShellFolder(0x%08x)::MapColumnToSCID(iColumn=%d)\n", this, iColumn);
#if defined(OW_PKEYS_SUPPORT)
	// This will map the columns to some built-in properties on Vista.
	// It's needed for the tile subtitles to display properly.
	switch (iColumn)
	{
	case DETAILS_COLUMN_NAME:
		*pscid = PKEY_ItemNameDisplay;
		return S_OK;
	case DETAILS_COLUMN_PATH:
		*pscid = PKEY_ItemPathDisplay;
		return S_OK;
	// We can seemingly skip rank and let it fall through to the legacy impl
	}
	return E_FAIL;
#endif
	return E_NOTIMPL;
}
