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

// stdafx.h : include file for standard system include files,
// or project specific include files that are used frequently, but
// are changed infrequently
//

#pragma once

// Property keys on Windows Vista, used for tile view subtitles.
// On older platforms, this can usually be still defined, but polyfilled.
// XXX: Perhaps it could disable all of IShellFolder2 for really old SDKs?
#define OW_PKEYS_SUPPORT

#include "targetver.h"

#define WIN32_LEAN_AND_MEAN             // Exclude rarely-used stuff from Windows headers
// Windows Header Files:
#include <windows.h>
// Must be before the other shell headers when we have mean and lean
#include <ShellAPI.h>

#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS      // some CString constructors will be explicit

//#define _ATL_DEBUG_QI

#include <atlbase.h>

extern CComModule _Module;

#include "wtlstr.h"
#include <atlcom.h>
#include <atlwin.h>

#include <ShlDisp.h>
#include <ShlObj.h>
#if defined(OW_PKEYS_SUPPORT) && VER_PRODUCTBUILD >= 6000
// These are probably supportable on 2600 in some obscure WDS SDK
#include <propkey.h>
#include <propvarutil.h>
#elif defined(OW_PKEYS_SUPPORT)
// Create polyfills for these (which don't link to anything, are just header
// magic) definitions. The DEFINE one is typedef' PROPERTYKEY in the real SDK,
// but we'll just save a line and define it to SHCOLUMNID for those old SDKs.
// These are taken from the Windows 7 Platform SDK.
#define DEFINE_PROPERTYKEY(name, l, w1, w2, b1, b2, b3, b4, b5, b6, b7, b8, pid) EXTERN_C const SHCOLUMNID DECLSPEC_SELECTANY name = { { l, w1, w2, { b1, b2,  b3,  b4,  b5,  b6,  b7,  b8 } }, pid }

DEFINE_PROPERTYKEY(PKEY_PropList_TileInfo, 0xC9944A21, 0xA406, 0x48FE, 0x82, 0x25, 0xAE, 0xC7, 0xE2, 0x4C, 0x21, 0x1B, 3);
DEFINE_PROPERTYKEY(PKEY_PropList_ExtendedTileInfo, 0xC9944A21, 0xA406, 0x48FE, 0x82, 0x25, 0xAE, 0xC7, 0xE2, 0x4C, 0x21, 0x1B, 9);
DEFINE_PROPERTYKEY(PKEY_PropList_PreviewDetails, 0xC9944A21, 0xA406, 0x48FE, 0x82, 0x25, 0xAE, 0xC7, 0xE2, 0x4C, 0x21, 0x1B, 8);
DEFINE_PROPERTYKEY(PKEY_PropList_FullDetails, 0xC9944A21, 0xA406, 0x48FE, 0x82, 0x25, 0xAE, 0xC7, 0xE2, 0x4C, 0x21, 0x1B, 2);
DEFINE_PROPERTYKEY(PKEY_ItemType, 0x28636AA6, 0x953D, 0x11D2, 0xB5, 0xD6, 0x00, 0xC0, 0x4F, 0xD9, 0x18, 0xD0, 11);
DEFINE_PROPERTYKEY(PKEY_ItemNameDisplay, 0xB725F130, 0x47EF, 0x101A, 0xA5, 0xF1, 0x02, 0x60, 0x8C, 0x9E, 0xEB, 0xAC, 10);
DEFINE_PROPERTYKEY(PKEY_ItemPathDisplay, 0xE3E0584C, 0xB788, 0x4A5A, 0xBB, 0x20, 0x7F, 0x5A, 0x44, 0xC9, 0xAC, 0xDD, 7);

#define IsEqualPropertyKey(a, b)   (((a).pid == (b).pid) && IsEqualIID((a).fmtid, (b).fmtid) )

// no in/out for VC++6 compat
inline HRESULT InitVariantFromString(PCWSTR psz, VARIANT *pvar)
{
    pvar->vt = VT_BSTR;
    pvar->bstrVal = SysAllocString(psz);
    HRESULT hr =  pvar->bstrVal ? S_OK : (psz ? E_OUTOFMEMORY : E_INVALIDARG);
    if (FAILED(hr))
    {
        VariantInit(pvar);
    }
    return hr;
}
#endif

#include <Shlwapi.h>
#include <UrlMon.h>
#include <WinInet.h>

