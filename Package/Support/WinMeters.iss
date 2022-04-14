; 'R:\!_Work\Alex_K\Progs\WinMeters\Package\SETUP.LST' imported by ISTool version 5.3.0

#define ApplicationName 'WinMeters'
#define ApplicationVersion GetFileVersion('WinMeters.exe')

[Setup]
AppName={#ApplicationName}
AppVerName={#ApplicationName} {#ApplicationVersion}
VersionInfoVersion={#ApplicationVersion}

;AppName=WinMeters
;AppVerName=WinMeters
AppContact=alex.commandor@gmail.com
DefaultDirName={pf}\WinMeters
DefaultGroupName=WinMeters
OutputBaseFilename=WinMeters_setup
OutputDir=.
VersionInfoCompany=alex.commandor@gmail.com
VersionInfoDescription=System activity meter installer
VersionInfoCopyright=2013 © Alex Commandor
VersionInfoProductName=WinMeters
MinVersion=0,5.01.2600
ShowLanguageDialog=no
DisableStartupPrompt=false
SetupIconFile=R:\!_Work\Alex_K\Progs\WinMeters\WinMeters_installer.ico

[Files]
; [Bootstrap Files]
; @COMCAT.DLL,$(WinSysPathSysFile),$(DLLSelfRegister),,5/31/98 12:00:00 AM,22288,4.71.1460.1
Source: COMCAT.DLL; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver noregerror
; @STDOLE2.TLB,$(WinSysPathSysFile),$(TLBRegister),,6/3/99 12:00:00 AM,17920,2.40.4275.1
Source: STDOLE2.TLB; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regtypelib noregerror onlyifdoesntexist
; @ASYCFILT.DLL,$(WinSysPathSysFile),,,3/8/99 12:00:00 AM,147728,2.40.4275.1
Source: ASYCFILT.DLL; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile
; @OLEPRO32.DLL,$(WinSysPathSysFile),$(DLLSelfRegister),,3/8/99 12:00:00 AM,164112,5.0.4275.1
Source: OLEPRO32.DLL; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver noregerror
; @OLEAUT32.DLL,$(WinSysPathSysFile),$(DLLSelfRegister),,4/12/00 12:00:00 AM,598288,2.40.4275.1
Source: OLEAUT32.DLL; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver noregerror
; @msvbvm60.dll,$(WinSysPathSysFile),$(DLLSelfRegister),,4/14/08 2:00:00 PM,1384479,6.0.98.2
Source: msvbvm60.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver noregerror

; [Setup1 Files]
; @MSCOMCTL.OCX,$(WinSysPath),$(DLLSelfRegister),$(Shared),5/2/12 12:17:12 PM,1070152,6.1.98.34
Source: MSCOMCTL.OCX; DestDir: {sys}; Flags: regserver sharedfile noregerror restartreplace uninsneveruninstall onlyifdoesntexist
; @msimg32.dll,$(WinSysPath),,$(Shared),2/18/07 7:00:00 AM,4608,5.2.3790.0
Source: msimg32.dll; DestDir: {sys}; Flags: sharedfile restartreplace onlyifdoesntexist
; @WinMeters.exe,$(AppPath),,,6/19/13 5:48:36 PM,1523712,1.0.0.88
Source: WinMeters.exe; DestDir: {app}; Flags: promptifolder restartreplace

[Icons]
Name: {group}\WinMeters; Filename: {app}\WinMeters.exe; WorkingDir: {app}; MinVersion: 0,5.01.2600
[Run]
Filename: {app}\WinMeters.exe; Flags: postinstall runascurrentuser nowait
