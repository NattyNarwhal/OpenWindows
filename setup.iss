[Setup]
AppName=Open Windows
AppCopyright=Copyright (c) 2020 Calvin Buckley
AppPublisher=Calvin Buckley
AppPublisherURL=https://github.com/NattyNarwhal/OpenWindows
AppVersion=1.4.0.1
DefaultDirName={pf}\Open Windows
DisableProgramGroupPage=yes
DisableReadyMemo=yes
DisableReadyPage=yes
; if we decide to support itanium, check this
ArchitecturesInstallIn64BitMode=x64

[Files]
Source: "OpenWindows-vc6-x86.dll"; DestDir: "{app}"; DestName: "OpenWindows.dll"; Flags: regserver 32bit; Check: not Is64BitInstallMode
; does this need regtypelib flag?
Source: "OpenWindows-vc6-x86.tlb"; DestDir: "{app}"; DestName: "OpenWindows.tlb"; Flags: 32bit; Check: not Is64BitInstallMode
Source: "OpenWindows-vc2010-amd64.dll"; DestDir: "{app}"; DestName: "OpenWindows.dll"; Flags: regserver 64bit; Check: Is64BitInstallMode
Source: "OpenWindows-vc2010-amd64.tlb"; DestDir: "{app}"; DestName: "OpenWindows.tlb"; Flags: 64bit; Check: Is64BitInstallMode

