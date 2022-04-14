; 'R:\!_Work\Alex_K\Progs\PhotoBookZakaz_v4\Package\Support\Setup.Lst' imported by ISTool version 5.3.0

#define ApplicationName 'Univest Digital projects'
#define ApplicationVersion GetFileVersion('UnivestDigital_v3.exe')

[Setup]
AppName={#ApplicationName}
AppVerName={#ApplicationName} {#ApplicationVersion}
VersionInfoVersion={#ApplicationVersion}
DefaultDirName={pf}\Univest Digital projects
DefaultGroupName=Univest Digital projects v3
OutputBaseFilename=UnivestDigital_v3_setup
VersionInfoCompany=Univest PrePress
VersionInfoCopyright=(c) 2013 Alex Commandor
ShowLanguageDialog=auto
[Files]
; [Bootstrap Files]
; @COMCAT.DLL,$(WinSysPathSysFile),$(DLLSelfRegister),,5/30/98 11:00:00 PM,22288,4.71.1460.1
Source: COMCAT.DLL; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver
; @STDOLE2.TLB,$(WinSysPathSysFile),$(TLBRegister),,6/2/99 11:00:00 PM,17920,2.40.4275.1
Source: STDOLE2.TLB; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regtypelib
; @ASYCFILT.DLL,$(WinSysPathSysFile),,,3/7/99 11:00:00 PM,147728,2.40.4275.1
Source: ASYCFILT.DLL; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile
; @OLEPRO32.DLL,$(WinSysPathSysFile),$(DLLSelfRegister),,3/7/99 11:00:00 PM,164112,5.0.4275.1
Source: OLEPRO32.DLL; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver
; @OLEAUT32.DLL,$(WinSysPathSysFile),$(DLLSelfRegister),,4/11/00 11:00:00 PM,598288,2.40.4275.1
Source: OLEAUT32.DLL; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver
; @msvbvm60.dll,$(WinSysPathSysFile),$(DLLSelfRegister),,4/14/08 1:00:00 PM,1384479,6.0.98.2
Source: msvbvm60.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver

; [Setup1 Files]
; @GflAx.dll,$(WinSysPath),$(DLLSelfRegister),$(Shared),2/27/08 4:32:14 PM,1167360,2.82.0.0
Source: GflAx.dll; DestDir: {sys}; Flags: promptifolder regserver sharedfile
; @MSCOMCTL.OCX,$(WinSysPath),$(DLLSelfRegister),$(Shared),5/2/12 11:17:12 AM,1070152,6.1.98.34
Source: MSCOMCTL.OCX; DestDir: {sys}; Flags: promptifolder regserver sharedfile uninsneveruninstall onlyifdoesntexist
; @UnivestDigital_v3.exe,$(AppPath),,,11/28/13 6:44:11 PM,1933312,3.1.0.28
Source: UnivestDigital_v3.exe; DestDir: {app}; Flags: promptifolder restartreplace

[Icons]
Name: {group}\Univest Digital projects v3; Filename: {app}\UnivestDigital_v3.exe; WorkingDir: {app}
[Run]
Filename: {app}\UnivestDigital_v3.exe; Flags: postinstall unchecked nowait
