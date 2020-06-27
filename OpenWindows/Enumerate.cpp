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
	CString appNameNormalized = CString(appName).MakeUpper();
	BOOL res = appNameNormalized.Find(_T("EXPLORER.EXE")) > -1;
	SysFreeString(appName);
	return res;
}

CString UriToDosPath(TCHAR *uri)
{
	DWORD size = INTERNET_MAX_URL_LENGTH;
	CString str;
	/* XXX: PathCreateFromUrl would properly accept a TCHAR* */
	CoInternetParseUrl(uri, PARSE_PATH_FROM_URL, 0, str.GetBuffer(size), size, &size, 0);
	str.ReleaseBuffer();
	return str;
}

CString SimplifyName(CString path)
{
	PathStripPath(path.GetBuffer(260));
	path.ReleaseBuffer();
	return path;
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

/* TODO: Convert to ATL wrappers */
long EnumerateExplorerWindows(COWItemList *list, HWND callerWindow)
{
	IShellWindows *windows;
	long count, realCount, i;
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
		
		BSTR locationUrl;
		CString str;
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

		// XXX: Do we use the ShellFolderView in get_Document?
		if (FAILED(wba->get_LocationURL(&locationUrl))) {
			ATLTRACE(_T(" ** Enumerate can't get location for i=%ld"), i);
			goto fail2;
		}

		str = UriToDosPath(locationUrl);
		if (str.GetLength() == 0) {
			ATLTRACE(_T(" ** Enumerate empty path string i=%ld"), i);
			goto fail3;
		}
		else if (str == PhysicalManifestationPath()) {
			// I hate this workaround around a workaround. The manifestation
			// path is used to give a (fake) real FS location for programs silly
			// enough to require one. This means if you have multiple of our NSE
			// though, you get ugly "Temp/" entries. Skip them if we encounter one.
			ATLTRACE(_T(" ** Enumerate path is the manifestation path i=%ld"), i);
			goto fail3;
		}

		ATLTRACE(_T(" ** Enumerate caught i=%ld # %ld: %s"), i, realCount, str);
		item.SetRank(realCount++);
		item.SetName(SimplifyName(str));
		item.SetPath(str);
		list->Add(item);

fail3:
		SysFreeString(locationUrl);
fail2:
		wba->Release();
fail1:
		wba_disp->Release();
	}
	windows->Release();
	return realCount;
}
