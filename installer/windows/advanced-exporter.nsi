; Advanced Exporter for PowerPoint: a per-user Windows installer.
;
; It copies manifest.xml into the user's own Programs folder and registers it
; with Office under HKCU\Software\Microsoft\Office\16.0\Wef\Developer: a
; string value named after the add-in's <Id>, holding the manifest's path.
; That is exactly what Microsoft's office-addin-dev-settings writes to sideload
; an add-in into desktop Office on Windows (read from its 4.0.1 source,
; 2026-10-24), and Office reads it at launch. No admin rights, no network-share
; catalog, nothing in Program Files.
;
; Built by installer/windows/build-exe.sh, which passes VERSION, ADDIN_ID,
; MANIFEST and OUTFILE. "16.0" covers Office 2016, 2019, 2021, 2024 and
; Microsoft 365.

Unicode true
ManifestDPIAware true
SetCompressor /SOLID lzma

!include "MUI2.nsh"
!include "LogicLib.nsh"

!ifndef VERSION
  !error "build with installer/windows/build-exe.sh"
!endif
!ifndef ADDIN_ID
  !error "build with installer/windows/build-exe.sh"
!endif
!ifndef MANIFEST
  !error "build with installer/windows/build-exe.sh"
!endif
!ifndef OUTFILE
  !error "build with installer/windows/build-exe.sh"
!endif

!define PRODUCT "Advanced Exporter for PowerPoint"
!define PUBLISHER "Stoatworks Labs"
!define SITE "https://pptx-exporter-addin.stoatworks-labs.com/"
!define WEF_KEY "Software\Microsoft\Office\16.0\Wef\Developer"
!define UNINST_KEY "Software\Microsoft\Windows\CurrentVersion\Uninstall\StoatworksAdvancedExporterForPowerPoint"

Name "${PRODUCT}"
OutFile "${OUTFILE}"
RequestExecutionLevel user
InstallDir "$LOCALAPPDATA\Programs\Stoatworks\Advanced Exporter for PowerPoint"
BrandingText "${PUBLISHER}"

VIProductVersion "${VERSION}.0"
VIAddVersionKey "ProductName" "${PRODUCT}"
VIAddVersionKey "CompanyName" "${PUBLISHER}"
VIAddVersionKey "FileDescription" "${PRODUCT} installer"
VIAddVersionKey "FileVersion" "${VERSION}"
VIAddVersionKey "ProductVersion" "${VERSION}"
VIAddVersionKey "LegalCopyright" "Copyright (c) 2026 Stoatworks Labs. MIT licence."

!define MUI_ICON "icon.ico"
!define MUI_UNICON "icon.ico"
!define MUI_ABORTWARNING

!define MUI_WELCOMEPAGE_TITLE "Advanced Exporter for PowerPoint"
!define MUI_WELCOMEPAGE_TEXT "This adds Advanced Exporter to PowerPoint for Windows, for your user account. It exports the open deck to the formats PowerPoint offers, with the options left in, and to the ones it never got round to: pictures at any size, handouts and notes pages, video, an HTML5 presentation, LaTeX, Markdown and a few more.$\r$\n$\r$\nThe installer puts one small file in your user folder and tells PowerPoint where it is. The add-in itself runs from pptx-exporter-addin.stoatworks-labs.com, so it keeps itself up to date, and your decks are never uploaded. No administrator rights are needed.$\r$\n$\r$\nPowerPoint has to be closed while it installs."
!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_INSTFILES
!define MUI_FINISHPAGE_TITLE "Advanced Exporter is in PowerPoint"
!define MUI_FINISHPAGE_TEXT "Open PowerPoint and any presentation. Advanced Exporter has a button on the Home tab; if it isn't there, choose Home > Add-ins and pick Advanced Exporter from your add-ins.$\r$\n$\r$\nTo remove it, uninstall Advanced Exporter for PowerPoint from Settings > Apps > Installed apps."
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES

!insertmacro MUI_LANGUAGE "English"

; PowerPoint reads the Developer key when it starts, so it has to be closed for
; an install or an uninstall to take effect. tasklist is on every Windows, and
; nsExec ships with NSIS, so this needs no plugin.
!macro WaitForPowerPointToClose UN
Function ${UN}WaitForPowerPointToClose
  check:
    nsExec::ExecToStack 'tasklist /FI "IMAGENAME eq POWERPNT.EXE" /NH /FO CSV'
    Pop $0
    Pop $1
    ${If} $0 != 0
      Return ; tasklist itself failed: don't trap anyone over a check
    ${EndIf}
    ; Running: a CSV row naming POWERPNT.EXE. Not running: an INFO line that
    ; doesn't mention it. A plain substring search, because StrFunc.nsh's
    ; StrStr crashed Homebrew's makensis 3.12 (bad_alloc, 2026-10-24).
    StrCpy $3 0
    search:
      StrCpy $2 $1 12 $3
      ${If} $2 == ""
        Return ; reached the end: PowerPoint isn't running
      ${EndIf}
      ${If} $2 == "POWERPNT.EXE"
        Goto running
      ${EndIf}
      IntOp $3 $3 + 1
      Goto search
  running:
    ${If} ${Silent}
      Abort ; a deployment tool gets a failure, not a dialog nobody will see
    ${EndIf}
    MessageBox MB_RETRYCANCEL|MB_ICONEXCLAMATION "PowerPoint is open. Save your work, close PowerPoint, then choose Retry." IDRETRY check
    Abort
FunctionEnd
!macroend
!insertmacro WaitForPowerPointToClose ""
!insertmacro WaitForPowerPointToClose "un."
Function .onInit
  Call WaitForPowerPointToClose
FunctionEnd

Function un.onInit
  Call un.WaitForPowerPointToClose
FunctionEnd

Section "Install"
  SetOutPath "$INSTDIR"
  File "/oname=manifest.xml" "${MANIFEST}"
  File "icon.ico"

  WriteRegStr HKCU "${WEF_KEY}" "${ADDIN_ID}" "$INSTDIR\manifest.xml"

  WriteUninstaller "$INSTDIR\Uninstall.exe"
  WriteRegStr HKCU "${UNINST_KEY}" "DisplayName" "${PRODUCT}"
  WriteRegStr HKCU "${UNINST_KEY}" "DisplayVersion" "${VERSION}"
  WriteRegStr HKCU "${UNINST_KEY}" "Publisher" "${PUBLISHER}"
  WriteRegStr HKCU "${UNINST_KEY}" "DisplayIcon" "$INSTDIR\icon.ico"
  WriteRegStr HKCU "${UNINST_KEY}" "InstallLocation" "$INSTDIR"
  WriteRegStr HKCU "${UNINST_KEY}" "URLInfoAbout" "${SITE}"
  WriteRegStr HKCU "${UNINST_KEY}" "UninstallString" '"$INSTDIR\Uninstall.exe"'
  WriteRegStr HKCU "${UNINST_KEY}" "QuietUninstallString" '"$INSTDIR\Uninstall.exe" /S'
  WriteRegDWORD HKCU "${UNINST_KEY}" "NoModify" 1
  WriteRegDWORD HKCU "${UNINST_KEY}" "NoRepair" 1
  WriteRegDWORD HKCU "${UNINST_KEY}" "EstimatedSize" 200
SectionEnd

Section "Uninstall"
  DeleteRegValue HKCU "${WEF_KEY}" "${ADDIN_ID}"
  ; Office's developer settings for this add-in (debugging flags), if anyone
  ; set them. Only this add-in's own subkey.
  DeleteRegKey HKCU "${WEF_KEY}\${ADDIN_ID}"

  Delete "$INSTDIR\manifest.xml"
  Delete "$INSTDIR\icon.ico"
  Delete "$INSTDIR\Uninstall.exe"
  RMDir "$INSTDIR"
  RMDir "$LOCALAPPDATA\Programs\Stoatworks" ; only goes if Font Manager or another of ours isn't in it

  DeleteRegKey HKCU "${UNINST_KEY}"
SectionEnd
