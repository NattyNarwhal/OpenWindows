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
#include "ShellItems.h"

//========================================================================================
// Helper for STRRET

CMalloc::CMalloc()
{
	HRESULT hr = SHGetMalloc(&m_MallocPtr);
	ATLASSERT(SUCCEEDED(hr));
}

CMalloc g_Malloc;

bool SetReturnStringA(LPCSTR Source, STRRET &str)
{
	ULONG StringLen = strlen(Source)+1;
	str.uType = STRRET_WSTR;
	str.pOleStr = (LPOLESTR)g_Malloc.m_MallocPtr->Alloc(StringLen*sizeof(OLECHAR));
	if (!str.pOleStr)
		return false;

	mbstowcs(str.pOleStr, Source, StringLen);
	return true;
}

bool SetReturnStringW(LPCWSTR Source, STRRET &str)
{
	ULONG StringLen = wcslen(Source)+1;
	str.uType = STRRET_WSTR;
	str.pOleStr = (LPOLESTR)g_Malloc.m_MallocPtr->Alloc(StringLen*sizeof(OLECHAR));
	if (!str.pOleStr)
		return false;

	wcsncpy(str.pOleStr, Source, StringLen);
	return true;
}

ULONG COWItem::GetSize()
{
	return   4
		   + 2
		   + 2
		   + (_tcslen(m_Path)+1)*sizeof(OLECHAR)
		   + (_tcslen(m_Name)+1)*sizeof(OLECHAR)
		   ;
}

void COWItem::CopyTo(void *pTarget)
{
	*(DWORD*)pTarget = MAGIC;
	*((USHORT*)pTarget+2) = 0;
	*((USHORT*)pTarget+3) = m_Rank;
#ifdef _UNICODE
	wcscpy((OLECHAR*)pTarget+4, m_Path);
	wcscpy((OLECHAR*)pTarget+4+wcslen(m_Path)+1, m_Name);
#else
	mbstowcs((OLECHAR*)pTarget+4, m_Path, strlen(m_Path)+1);
	mbstowcs((OLECHAR*)pTarget+4+strlen(m_Path)+1, m_Name, strlen(m_Name)+1);
#endif

}

//-------------------------------------------------------------------------------

void COWItem::SetPath(LPCTSTR Path)
{
	m_Path = Path;
}

void COWItem::SetName(LPCTSTR Name)
{
	m_Name = Name;
}

void COWItem::SetRank(USHORT Rank)
{
	m_Rank = Rank;
}

//-------------------------------------------------------------------------------

bool COWItem::IsOwn(LPCITEMIDLIST pidl)
{
	if ((pidl == NULL) || (pidl->mkid.cb < 4))
		return false;

	return *((DWORD*)(pidl->mkid.abID)) == MAGIC;
}

LPOLESTR COWItem::GetPath(LPCITEMIDLIST pidl)
{
	return (OLECHAR*)pidl+5;
}

LPOLESTR COWItem::GetName(LPCITEMIDLIST pidl)
{
	return (OLECHAR*) ( (BYTE*)pidl + 10 + (wcslen((OLECHAR*)pidl+5)+1) * sizeof(OLECHAR) );
}

USHORT COWItem::GetRank(LPCITEMIDLIST pidl)
{
	return *((USHORT*)pidl+4);
}

//========================================================================================
// CDataObject

// helper function that creates a CFSTR_SHELLIDLIST format from given pidls.
HGLOBAL CreateShellIDList(LPCITEMIDLIST pidlParent, LPCITEMIDLIST *aPidls, UINT uItemCount)
{
	HGLOBAL hGlobal = NULL;
	LPIDA pData;
	UINT iCurPos;
	UINT cbPidl;
	UINT i;
	CPidlMgr PidlMgr;

	// get the size of the parent folder's PIDL
	cbPidl = PidlMgr.GetSize(pidlParent);

	// get the total size of all of the PIDLs
	for (i=0; i<uItemCount; i++)
	{
		cbPidl += PidlMgr.GetSize(aPidls[i]);
	}

	// Find the end of the CIDA structure. This is the size of the 
	// CIDA structure itself (which includes one element of aoffset) plus the 
	// additional number of elements in aoffset.
	iCurPos = sizeof(CIDA) + (uItemCount * sizeof(UINT));

	// Allocate the memory for the CIDA structure and it's variable length members.
	hGlobal = GlobalAlloc(GPTR | GMEM_SHARE, (DWORD)
	                                         (iCurPos +        // size of the CIDA structure and the additional aoffset elements
	                                         (cbPidl + 1)));   // size of the pidls
	if (!hGlobal)
		return NULL;

	if (pData = (LPIDA)GlobalLock(hGlobal))
	{
		pData->cidl = uItemCount;
		pData->aoffset[0] = iCurPos;

		// add the PIDL for the parent folder
		cbPidl = PidlMgr.GetSize(pidlParent);
		CopyMemory((LPBYTE)(pData) + iCurPos, (LPBYTE)pidlParent, cbPidl);
		iCurPos += cbPidl;

		for (i=0; i<uItemCount; i++)
		{
			// get the size of the PIDL
			cbPidl = PidlMgr.GetSize(aPidls[i]);

			// fill out the members of the CIDA structure.
			pData->aoffset[i + 1] = iCurPos;

			// copy the contents of the PIDL
			CopyMemory((LPBYTE)(pData) + iCurPos, (LPBYTE)aPidls[i], cbPidl);

			// set up the position of the next PIDL
			iCurPos += cbPidl;
		}

		GlobalUnlock(hGlobal);
	}

	return hGlobal;
}

CDataObject::CDataObject() : m_pidl(NULL), m_pidlParent(NULL)
{
	m_cfShellIDList = RegisterClipboardFormat(CFSTR_SHELLIDLIST);
}

CDataObject::~CDataObject()
{
	m_PidlMgr.Delete(m_pidl);
	m_PidlMgr.Delete(m_pidlParent);
}

void CDataObject::Init(IUnknown *pUnkOwner)
{
	m_UnkOwnerPtr = pUnkOwner;
}

void CDataObject::SetPidl(LPCITEMIDLIST pidlParent, LPCITEMIDLIST pidl)
{
	m_pidlParent = m_PidlMgr.Copy(pidlParent);
	m_pidl = m_PidlMgr.Copy(pidl);
}

//-------------------------------------------------------------------------------

STDMETHODIMP CDataObject::GetData(LPFORMATETC pFE, LPSTGMEDIUM pStgMedium)
{
	ATLTRACE("CDataObject::GetData()\n");

	if (pFE->cfFormat == m_cfShellIDList)
	{
		pStgMedium->hGlobal = CreateShellIDList(m_pidlParent, (LPCITEMIDLIST*)&m_pidl, 1);

		if (pStgMedium->hGlobal)
		{
			pStgMedium->tymed = TYMED_HGLOBAL;
			pStgMedium->pUnkForRelease = NULL;	// Even if our tymed is HGLOBAL, WinXP calls ReleaseStgMedium() which tries to call pUnkForRelease->Release() : BANG!
			return S_OK;
		}
	}

	return E_INVALIDARG;
}

STDMETHODIMP CDataObject::GetDataHere(LPFORMATETC, LPSTGMEDIUM)
{
	ATLTRACE("CDataObject::GetDataHere()\n");
	return E_NOTIMPL;
}

STDMETHODIMP CDataObject::QueryGetData(LPFORMATETC)
{
	ATLTRACE("CDataObject::QueryGetData()\n");
	return E_NOTIMPL;
}

STDMETHODIMP CDataObject::GetCanonicalFormatEtc(LPFORMATETC, LPFORMATETC)
{
	ATLTRACE("CDataObject::GetCanonicalFormatEtc()\n");
	return E_NOTIMPL;
}

STDMETHODIMP CDataObject::SetData(LPFORMATETC, LPSTGMEDIUM, BOOL)
{
	ATLTRACE("CDataObject::SetData()\n");
	return E_NOTIMPL;
}

STDMETHODIMP CDataObject::EnumFormatEtc(DWORD, IEnumFORMATETC**)
{
	ATLTRACE("CDataObject::EnumFormatEtc()\n");
	return E_NOTIMPL;
}

STDMETHODIMP CDataObject::DAdvise(LPFORMATETC, DWORD, IAdviseSink*, LPDWORD)
{
	ATLTRACE("CDataObject::DAdvise()\n");
	return E_NOTIMPL;
}

STDMETHODIMP CDataObject::DUnadvise(DWORD dwConnection)
{
	ATLTRACE("CDataObject::DUnadvise()\n");
	return E_NOTIMPL;
}

STDMETHODIMP CDataObject::EnumDAdvise(IEnumSTATDATA** ppEnumAdvise)
{
	ATLTRACE("CDataObject::EnumDAdvise()\n");
	return E_NOTIMPL;
}

//-------------------------------------------------------------------------------

STDMETHODIMP CDataObject::Next(ULONG, LPFORMATETC, ULONG*)
{
	ATLTRACE("CDataObject::Next()\n");
	return E_NOTIMPL;
}

STDMETHODIMP CDataObject::Skip(ULONG)
{
	ATLTRACE("CDataObject::Skip()\n");
	return E_NOTIMPL;
}

STDMETHODIMP CDataObject::Reset()
{
	ATLTRACE("CDataObject::Reset()\n");
	return E_NOTIMPL;
}

STDMETHODIMP CDataObject::Clone(LPENUMFORMATETC*)
{
	ATLTRACE("CDataObject::Clone()\n");
	return E_NOTIMPL;
}


