#pragma once

// Including SDKDDKVer.h defines the highest available Windows platform.

// If you wish to build your application for a previous Windows platform, include WinSDKVer.h and
// set the _WIN32_WINNT macro to the platform you wish to support before including SDKDDKVer.h.

// For getting the version of the SDK regardless of what we want
#include <ntverp.h>
// When did the SDK get this? I'm assuming Vista, but...
#if VER_PRODUCTBUILD >= 6000
#include <SDKDDKVer.h>
#else
// We'll just assume Windows 2000 minreq to satisify ATL, but it should work
// on NT 4 and 9x as well. (Maybe 0x0403 is a safer compromise?)
#define _WIN32_WINNT 0x0500
#define _WIN32_IE 0x0500
#endif
