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


#ifndef __ROOTSHELLFOLDER_H_
#define __ROOTSHELLFOLDER_H_

#include "resource.h"       // main symbols

#include "MPidlMgr.h"
using namespace Mortimer;

#include "ShellItems.h"
#include "Enumerate.h"

#include "CComEnumOnCArray.h"

//========================================================================================

enum
{
	DETAILS_COLUMN_NAME,
	DETAILS_COLUMN_PATH,
	DETAILS_COLUMN_RANK,

	DETAILS_COLUMN_MAX
};

//========================================================================================
// COWRootShellFolder

class ATL_NO_VTABLE COWRootShellFolder : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<COWRootShellFolder, &CLSID_OpenWindowsRootShellFolder>,
	public IShellFolder2,
    public IPersistFolder2,
	public IShellDetails
{
public:
	COWRootShellFolder();

DECLARE_REGISTRY_RESOURCEID(IDR_ROOTSHELLFOLDER)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(COWRootShellFolder)
	COM_INTERFACE_ENTRY(IShellFolder)
	COM_INTERFACE_ENTRY(IShellFolder2)
	COM_INTERFACE_ENTRY(IPersistFolder)
	COM_INTERFACE_ENTRY(IPersistFolder2)
	COM_INTERFACE_ENTRY(IPersist)
	COM_INTERFACE_ENTRY_IID(IID_IShellDetails, IShellDetails)
END_COM_MAP()

public:
	//-------------------------------------------------------------------------------
	// IPersist

	STDMETHOD(GetClassID)(CLSID*);

	//-------------------------------------------------------------------------------
	// IPersistFolder(2)

	STDMETHOD(Initialize)(LPCITEMIDLIST);
	STDMETHOD(GetCurFolder)(LPITEMIDLIST *ppidl);

	//-------------------------------------------------------------------------------
	// IShellFolder

	STDMETHOD(BindToObject) (LPCITEMIDLIST, LPBC, REFIID, void**);
	STDMETHOD(CompareIDs) (LPARAM, LPCITEMIDLIST, LPCITEMIDLIST);
	STDMETHOD(CreateViewObject) (HWND, REFIID, void** );
	STDMETHOD(EnumObjects) (HWND, DWORD, LPENUMIDLIST*);
	STDMETHOD(GetAttributesOf) (UINT, LPCITEMIDLIST*, LPDWORD);
	STDMETHOD(GetUIObjectOf) (HWND, UINT, LPCITEMIDLIST*, REFIID, LPUINT, void**);
	STDMETHOD(BindToStorage) (LPCITEMIDLIST, LPBC, REFIID, void**);
	STDMETHOD(GetDisplayNameOf) (LPCITEMIDLIST, DWORD uFlags, LPSTRRET lpName);
	STDMETHOD(ParseDisplayName) (HWND, LPBC, LPOLESTR, LPDWORD, LPITEMIDLIST*, LPDWORD);
	STDMETHOD(SetNameOf) (HWND, LPCITEMIDLIST, LPCOLESTR, DWORD, LPITEMIDLIST*);

	//-------------------------------------------------------------------------------
	// IShellDetails

	STDMETHOD(ColumnClick) (UINT iColumn);
	STDMETHOD(GetDetailsOf) (LPCITEMIDLIST pidl, UINT iColumn, LPSHELLDETAILS pDetails);

	//-------------------------------------------------------------------------------
	// IShellFolder2

	STDMETHOD(EnumSearches) (IEnumExtraSearch **ppEnum);
	STDMETHOD(GetDefaultColumn) (DWORD dwReserved, ULONG *pSort, ULONG *pDisplay);
	STDMETHOD(GetDefaultColumnState) (UINT iColumn, SHCOLSTATEF *pcsFlags);
	STDMETHOD(GetDefaultSearchGUID) (GUID *pguid);
	STDMETHOD(GetDetailsEx) (LPCITEMIDLIST pidl, const SHCOLUMNID *pscid, VARIANT *pv);
	// already in IShellDetails: STDMETHOD(GetDetailsOf) (LPCITEMIDLIST pidl, UINT iColumn, SHELLDETAILS *psd);
	STDMETHOD(MapColumnToSCID) (UINT iColumn, SHCOLUMNID *pscid);

	//-------------------------------------------------------------------------------

protected:
	CPidlMgr m_PidlMgr;

	LPITEMIDLIST m_pidlRoot;

	COWItemList m_OpenedWindows;
};

#endif //__ROOTSHELLFOLDER_H_
