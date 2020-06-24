# OpenWindows

This is a shell namespace extension that displays a list of opened Windows
Explorer windows you have open. It's inspired by the OS/2 stock file dialogs
and how they did the same for Workplace Shell windows. If you're the type of
person who has a lot of Explorer windows open, and wish you could get to one
from a save dialog, this is probably the extension for you.

Shell namespace extensions are one of the hairiest parts of the already pretty
baroque COM shell world due to the thin documentation on them and few debugger
aids availabile. Good examples are even thinner. Hopefully, this provides an
example of a shell extension that does something and has the least amount of
untangling layers.

## Building

This is a Visual C++ 2010 project. It'll build for x86 and amd64. It's been
tested on XP, Vista, and 10.

Building with newer Visual Studio should work fine. Building with older Visual
Studio would enable compatibility going back to 95/NT4 with the IE4 webby
shell, but I haven't had much luck getting the code to compile with VC++6 or
VC++2003. Patches welcome if you can figure that out; some is obvious like
handling PKEY/CString/IDL differences, others less so.

## Installation

Copy the DLL and TLB somehwere and run `regsvr32 OpenWindows.dll`.

To uninstall, run `regsvr32 /u OpenWindows.dll`.

This is going to worm its way into anything with a shell view, so it might
cause stability issues. Caveat emptor.

## Known issues

* The code isn't as 64-bit clean as it should be, and there are some unchecked
  string handling functions used (2004 was a more innocent time).
* The namespace alleges it has a physical manifestation. This is appearantly
  to appease software that has custom dialogs checking for filesystem presence
  such as... Microsoft Office. Oops! This is the system temp folder, so a
  window opened there won't appear. Sorry! (It could be mitigated if we pick
  an install path...)
* We could have an installer which would register it and create paths.
* Modern property handling is very, very basic and intended just to appease
  the tile view subtitles in Vista+. It should probably use fancier behaviour,
  but for now it delegates as much as it can except the bare minimum to the
  "classic" way of doing it.
* The icon sucks, doesn't it?
* There's probably a better way to do a lot of this, but again, it's a lot of
  cargo culting. Eventually I did find some wrapper for this... after I was
  well into the MVP. Mea culpa. (They prob wouldn't fit the needs very well,
  since I suspect most people interested in NSEs want to materialize entities
  that don't exist in the filesystem. Pascal's code did and provided a good
  start.)
* `Enumerate.cpp` started from a simple example, so it doesn't use ATL much.

## Attributions

See `LICENSE.txt`.

This is based off of code by Pascal Hurni. I contacted him for clarifying the
license of the derived code and he has approved of putting it under the MIT
license.

Source (a DOpus favourites NSE):
https://www.codeproject.com/Articles/7973/An-almost-complete-Namespace-Extension-Sample