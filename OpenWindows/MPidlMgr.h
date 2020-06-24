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
// Greatly inspired from several examples from: Microsoft, Michael Dunn, ...

#ifndef __MORTIMER_PIDLMGR_H__
#define __MORTIMER_PIDLMGR_H__

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

//========================================================================================
// encapsulate these classes in a namespace

#ifndef DOXYGEN_SHOULD_SKIP_THIS
namespace Mortimer
{
#endif

//========================================================================================

class CPidlData
{
public:
//	CPidlData();
//	CPidlData(const CPidlData &src);
//	virtual CPidlData &operator=(const CPidlData &src);

	virtual ULONG GetSize() = 0;
	virtual void CopyTo(void *pTarget) = 0;
};


class CPidlMgr
{
public:
	CPidlMgr()
	{
		HRESULT hr = SHGetMalloc(&m_MallocPtr);
		ATLASSERT(SUCCEEDED(hr));
	}

	LPITEMIDLIST Create(CPidlData &Data)
	{
		// Total size of the PIDL, including SHITEMID
		UINT TotalSize = sizeof(ITEMIDLIST) + Data.GetSize();

		// Also allocate memory for the final null SHITEMID.
		LPITEMIDLIST pidlNew = (LPITEMIDLIST) m_MallocPtr->Alloc(TotalSize + sizeof(ITEMIDLIST));
		if (pidlNew)
		{
			LPITEMIDLIST pidlTemp = pidlNew;

			// Prepares the PIDL to be filled with actual data
			pidlTemp->mkid.cb = TotalSize;

			// Fill the PIDL
			Data.CopyTo((void*)pidlTemp->mkid.abID);

			// Set an empty PIDL at the end
			pidlTemp = GetNextItem(pidlTemp);
			pidlTemp->mkid.cb = 0;
			pidlTemp->mkid.abID[0] = 0;
		}

		return pidlNew;
	}

	void Delete(LPITEMIDLIST pidl)
	{
		if (pidl)
			m_MallocPtr->Free(pidl);
	}

	LPITEMIDLIST GetNextItem(LPCITEMIDLIST pidl)
	{
		ATLASSERT(pidl != NULL);
		if (!pidl)
			return NULL;

		return LPITEMIDLIST(LPBYTE(pidl) + pidl->mkid.cb);
	}

	LPITEMIDLIST GetLastItem(LPCITEMIDLIST pidl)
	{
		LPITEMIDLIST pidlLast = NULL;

		//get the PIDL of the last item in the list
		while (pidl && pidl->mkid.cb)
		{
			pidlLast = (LPITEMIDLIST)pidl;
			pidl = GetNextItem(pidl);
		}

		return pidlLast;
	}

	LPITEMIDLIST Copy(LPCITEMIDLIST pidlSrc)
	{
		LPITEMIDLIST pidlTarget = NULL;
		UINT Size = 0;

		if (pidlSrc == NULL)
			return NULL;

		// Allocate memory for the new PIDL.
		Size = GetSize(pidlSrc);
		pidlTarget = (LPITEMIDLIST) m_MallocPtr->Alloc(Size);

		if (pidlTarget == NULL)
			return NULL;

		// Copy the source PIDL to the target PIDL.
		CopyMemory(pidlTarget, pidlSrc, Size);

		return pidlTarget;
	}

	UINT GetSize(LPCITEMIDLIST pidl)
	{
		UINT Size = 0;
		LPITEMIDLIST pidlTemp = (LPITEMIDLIST) pidl;

		ATLASSERT(pidl != NULL);
		if (!pidl)
			return 0;

		while (pidlTemp->mkid.cb != 0)
		{
			Size += pidlTemp->mkid.cb;
			pidlTemp = GetNextItem(pidlTemp);
		}  

		// add the size of the NULL terminating ITEMIDLIST
		Size += sizeof(ITEMIDLIST);

		return Size;
	}

	bool IsSingle(LPCITEMIDLIST pidl)
	{
		LPITEMIDLIST pidlTemp = GetNextItem(pidl);
		return pidlTemp->mkid.cb == 0;
	}

	CString StrRetToCString(STRRET *pStrRet, LPCITEMIDLIST pidl)
	{
		int Length = 0;
		bool Unicode = false;
		LPCSTR StringA = NULL;
		LPCWSTR StringW = NULL;

		switch (pStrRet->uType)
		{
		case STRRET_CSTR:
			StringA = pStrRet->cStr;
			Unicode = false;
			Length = strlen(StringA);
			break;

		case STRRET_OFFSET:
			StringA = (char*)pidl+pStrRet->uOffset;
			Unicode = false;
			Length = strlen(StringA);
			break;

		case STRRET_WSTR:
			StringW = pStrRet->pOleStr;
			Unicode = true;
			Length = wcslen(StringW);
			break;
		}

		if (Length==0)
			return CString();

		CString Target;
		LPTSTR pTarget = Target.GetBuffer(Length);
		if (Unicode)
		{
#ifdef _UNICODE
			wcscpy(pTarget, StringW);
#else
			wcstombs(pTarget, StringW, Length+1);
#endif
		}
		else
		{
#ifdef _UNICODE
			mbstowcs(pTarget, StringA, Length+1);
#else
			strcpy(pTarget, StringA);
#endif
		}
		Target.ReleaseBuffer();

		// Release the OLESTR
		if (pStrRet->uType == STRRET_WSTR)
		{
			m_MallocPtr->Free(pStrRet->pOleStr);
		}

		return Target;
	}

protected:
	CComPtr<IMalloc> m_MallocPtr;
};


//========================================================================================

}; // namespace Mortimer

#endif // __MORTIMER_PIDLMGR_H__
