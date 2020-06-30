/*
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

CString PhysicalManifestationPath(void)
{
	// Workaround. See COWRootShellFolder::GetDisplayNameOf.
	CString str;
	// yes, it really is the opposite
	GetTempPath(MAX_PATH, str.GetBuffer(MAX_PATH));
	str.ReleaseBuffer();
	return str;
}

BOOL IsExplorerWindow(IWebBrowserApp *wba)
{
	BSTR appName;
	wba->get_FullName(&appName);
	CString appNameNormalized = CString(appName);
	/* sep step because wtlstr doesn't return CString for mutators */
	appNameNormalized.MakeUpper();
	BOOL res = appNameNormalized.Find(_T("EXPLORER.EXE")) > -1;
	SysFreeString(appName);
	return res;
}

#if _DEBUG
void TraceHwndInner(HWND tracedWindow, TCHAR *desc)
{
	TCHAR name[255], klass[255];
	GetClassName(tracedWindow, klass, 255);
	GetWindowText(tracedWindow, name, 255);
	ATLTRACE(_T(" ** TraceHwnd %ld %s: name %s class %s"), (long)tracedWindow, desc, name, klass);
}

#define TraceHwnd(t, d) TraceHwndInner(t, d)
#else
#define TraceHwnd(x, y)
#endif

static BOOL FolderItemStrategy(IWebBrowserApp *wba, int i, BSTR *nameBStr, BSTR *pathBStr)
{
	IDispatch *sfvd_disp;
	IShellFolderViewDual *sfvd;
	Folder *folder;
	Folder2 *folder2;
	FolderItem *selfItem;
	BOOL ok;

	ok = TRUE;

	// This huge involved process boils down to:
	// - get the scriptable shell view from the browser object (thank IE)
	// - get the active folder from that, cast it into a newer interface
	// - get the folder as an item from the casted version
	// - get the name and path from the item
	// Unfortunately, this requires quite a bit of COM casting :/
	if (FAILED(wba->get_Document(&sfvd_disp))) {
		ATLTRACE(_T(" ** Enumerate can't get document dispatch for i=%ld"), i);
		ok = FALSE;
		goto fail2;
	}
	if (FAILED(sfvd_disp->QueryInterface(IID_IShellFolderViewDual, (void**)&sfvd))) {
		ATLTRACE(_T(" ** Enumerate isn't an IShellFolderViewDual i=%ld"), i);
		ok = FALSE;
		goto fail3;
	}
	if (FAILED(sfvd->get_Folder(&folder))) {
		ATLTRACE(_T(" ** Enumerate can't get folder i=%ld"), i);
		ok = FALSE;
		goto fail4;
	}
	if (FAILED(folder->QueryInterface(IID_Folder2, (void**)&folder2))) {
		ATLTRACE(_T(" ** Enumerate isn't a Folder2 i=%ld"), i);
		ok = FALSE;
		goto fail5;
	}
	// This part seems to fail on Me, possibly other 9x with 0xC0000005.
	if (FAILED(folder2->get_Self(&selfItem))) {
		ATLTRACE(_T(" ** Enumerate can't get FolderItem i=%ld"), i);
		ok = FALSE;
		goto fail6;
	}
	if (FAILED(selfItem->get_Name(nameBStr))) {
		ATLTRACE(_T(" ** Enumerate doesn't have folder name i=%ld"), i);
		ok = FALSE;
		goto fail7;
	}
	if (FAILED(selfItem->get_Path(pathBStr))) {
		ATLTRACE(_T(" ** Enumerate doesn't have folder path i=%ld"), i);
		ok = FALSE;
		goto fail8;
	}
fail8:
	SysFreeString(*nameBStr);
fail7:
	selfItem->Release();
fail6:
	folder2->Release();
fail5:
	folder->Release();
fail4:
	sfvd->Release();
fail3:
	sfvd_disp->Release();
fail2:
	return ok;
}

static BSTR UriToDosPath(BSTR uri)
{
	DWORD size = INTERNET_MAX_URL_LENGTH;
#pragma comment(lib, "urlmon")
	// Using the SHLWAPI function is tempting, but it's broken with fancy
	// characters even with Unicode
	wchar_t strW[INTERNET_MAX_URL_LENGTH];
	CoInternetParseUrl(uri, PARSE_PATH_FROM_URL, 0, strW, size, &size, 0);
	return SysAllocString(strW);
}

static BSTR SimplifyName(BSTR path)
{
	wchar_t newPath[MAX_PATH];
	wcsncpy(newPath, path, MAX_PATH);
#pragma comment(lib, "shlwapi")
	// Even 9x has the wide functions, because IE
	PathStripPathW(newPath);
	return SysAllocString(newPath);
}

static BOOL FileUriStrategy(IWebBrowserApp *wba, int i, BSTR *nameBStr, BSTR *pathBStr)
{
	BSTR locationUrl, name, path;
	BOOL ok;

	ok = TRUE;

	if (FAILED(wba->get_LocationURL(&locationUrl))) {
		ATLTRACE(_T(" ** Enumerate can't get location for i=%ld"), i);
		ok = FALSE;
		goto fail2;
	}

	path = UriToDosPath(locationUrl);
	name = SimplifyName(path);
	*pathBStr = path;
	*nameBStr = name;

	SysFreeString(locationUrl);
fail2:
	return ok;
}

/* TODO: Convert to ATL wrappers */
long EnumerateExplorerWindows(COWItemList *list, HWND callerWindow)
{
	IShellWindows *windows;
	long count, realCount, i;
	CString physPath;
	physPath = PhysicalManifestationPath();
	realCount = 0;
	if (FAILED(CoInitialize(NULL))) {
		ATLTRACE(_T(" ** Enumerate can't init COM"));
		return 1;
	}
	if (FAILED(CoCreateInstance(CLSID_ShellWindows, NULL, CLSCTX_ALL, IID_IShellWindows, (void**)&windows))) {
		ATLTRACE(_T(" ** Enumerate can't create IShellWindows"));
		return 2;
	}
	if (FAILED(windows->get_Count(&count))) {
		count = 0;
	}
	for (i = 0; i < count; i++) {
		VARIANT v ;
		v.vt = VT_I4 ;
		V_I4(&v) = i;

		IDispatch *wba_disp;
		IWebBrowserApp *wba;
		
		BSTR pathBStr, nameBStr;
		CString pathStr, nameStr;
		COWItem item;
		HWND window, parent;
		SHANDLE_PTR windowPtr;

		if (FAILED(windows->Item(v, &wba_disp))) {
			ATLTRACE(_T(" ** Enumerate isn't an item i=%ld"), i);
			continue;
		}
		if (FAILED(wba_disp->QueryInterface(IID_IWebBrowserApp, (void**)&wba))) {
			ATLTRACE(_T(" ** Enumerate isn't an IWebBrowserApp i=%ld"), i);
			goto fail1;
		}
		// Is this even a Windows Explorer window?
		if (!IsExplorerWindow(wba)) {
			ATLTRACE(_T(" ** Enumerate isn't an explorer window i=%ld"), i);
			goto fail1;
		}

		// Otherwise, we could have the window containing the the enumeration
		// be included. (It'll display the path of the previous folder, or
		// display this NSE when you refresh.)
		TraceHwnd(callerWindow, _T("caller"));
		if (FAILED(wba->get_HWND(&windowPtr))) {
			ATLTRACE(_T(" ** Enumerate failed to get the HWND for i=%ld"), i);
			goto fail1;
		}
		window = (HWND)windowPtr;
		TraceHwnd((HWND)window, _T("received"));

		ATLTRACE(_T(" ** Enumerate i=%ld callerHwnd=%ld vs. receivedHwnd %ld"),
			i, (long)callerWindow, window);
		if (callerWindow == window) {
			ATLTRACE(_T(" ** Enumerate windows are the same i=%ld"), i);
			goto fail1;
		}
		// On Vista, we don't get a CabinetWClass as the caller, but a
		// ShellTabWindowClass. Depending on if the sidebar or caller
		// window is fetching, the direct parent may or may not be the
		// expected CabinetNClass, so try to jump up to the top.
		else {
			parent = GetAncestor(callerWindow, GA_ROOTOWNER);
			TraceHwnd(parent, _T("parent of caller"));
			if (parent == window) {
				ATLTRACE(_T(" ** Enumerate windows are the same (checking parent of caller) i=%ld"), i);
				goto fail1;
			}
		}

		// Unfortunately, while the folder item strategy is preferred,
		// it has issues on Me. Fall back to the file:// URI strategy
		// if it fails.
		if (!FolderItemStrategy(wba, i, &nameBStr, &pathBStr)) {
			ATLTRACE(_T(" ** Enumerate folder item strat failed i=%ld"), i);
			if (!FileUriStrategy(wba, i, &nameBStr, &pathBStr)) {
				ATLTRACE(_T(" ** Enumerate file URI strat failed (bail) i=%ld"), i);
				goto fail2;
			}
		}

		nameStr = CString(nameBStr);
		pathStr = CString(pathBStr);
		if (pathStr.GetLength() == 0) {
			ATLTRACE(_T(" ** Enumerate empty path string i=%ld"), i);
			goto fail3;
		}
		else if (pathBStr[0] == L':' && pathBStr[1] == L':') {
			// This path is some shell namespace world stuff. This on its own
			// isn't inherently wrong, but it seems a bit random (or not, but
			// maybe just finicky about path syntax) if it'll actually point
			// to the object, or be inert. Unless we figure out a good way
			// to deal with this, for now, we can just ignore them.
			// (Or make it toggleable?)
			ATLTRACE(_T(" ** Enumerate skipping shell namespace i=%ld"), i);
			goto fail3;
		}
		else if (pathStr == physPath) {
			// I hate this workaround around a workaround. The manifestation
			// path is used to give a (fake) real FS location for programs silly
			// enough to require one. This means if you have multiple of our NSE
			// though, you get ugly "Temp/" entries. Skip them if we encounter one.
			ATLTRACE(_T(" ** Enumerate path is the manifestation path i=%ld"), i);
			goto fail3;
		}

		ATLTRACE(_T(" ** Enumerate caught i=%ld # %ld: %s <- %s"), i, realCount, nameStr, pathStr);
		item.SetRank(realCount++);
		item.SetName(nameBStr);
		item.SetPath(pathBStr);
		list->Add(item);

fail3:
		SysFreeString(nameBStr);
		SysFreeString(pathBStr);
fail2:
		wba->Release();
fail1:
		wba_disp->Release();
	}
	windows->Release();
	return realCount;
}
