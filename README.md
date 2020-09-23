<div align="center">

## A Way to Change The Resolution


</div>

### Description

This is ideal for a game. Many games simply look better or work better in a different resolution. Mostly the look is because whatever resolution you were using on your computer at the time is what the graphics in the game will look best as. This code will simply enable you to set their computer resolution to whatever you want.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Steve](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/steve.md)
**Level**          |Unknown
**User Rating**    |4.2 (164 globes from 39 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Games](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/games__1-38.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/steve-a-way-to-change-the-resolution__1-3840/archive/master.zip)

### API Declarations

```
'This is all the api stuff. It's not to difficult...
'check out my site at http://www.vbtutor.com
Declare Function EnumDisplaySettings Lib "user32" _
Alias "EnumDisplaySettingsA" _
(ByVal lpszDeviceName As Long, _
ByVal iModeNum As Long, _
lpDevMode As Any) As BooleanDeclare Function ChangeDisplaySettings Lib "user32" _
Alias "ChangeDisplaySettingsA" _
(lpDevMode As Any, ByVal dwFlags As Long) As Long
Declare Function ExitWindowsEx Lib "user32" _
(ByVal uFlags As Long, ByVal dwReserved As Long) As LongPublic Const EWX_LOGOFF = 0
Public Const EWX_SHUTDOWN = 1
Public Const EWX_REBOOT = 2
Public Const EWX_FORCE = 4
Public Const CCDEVICENAME = 32
Public Const CCFORMNAME = 32
Public Const DM_BITSPERPEL = &H40000
Public Const DM_PELSWIDTH = &H80000
Public Const DM_PELSHEIGHT = &H100000
Public Const CDS_UPDATEREGISTRY = &H1
Public Const CDS_TEST = &H4
Public Const DISP_CHANGE_SUCCESSFUL = 0
Public Const DISP_CHANGE_RESTART = 1Type DEVMODE
  dmDeviceName As String * CCDEVICENAME
  dmSpecVersion As Integer
  dmDriverVersion As Integer
  dmSize As Integer
  dmDriverExtra As Integer
  dmFields As Long
  dmOrientation As Integer
  dmPaperSize As Integer
  dmPaperLength As Integer
  dmPaperWidth As Integer
  dmScale As Integer
  dmCopies As Integer
  dmDefaultSource As Integer
  dmPrintQuality As Integer
  dmColor As Integer
  dmDuplex As Integer
  dmYResolution As Integer
  dmTTOption As Integer
  dmCollate As Integer
  dmFormName As String * CCFORMNAME
  dmUnusedPadding As Integer
  dmBitsPerPel As Integer
  dmPelsWidth As Long
  dmPelsHeight As Long
  dmDisplayFlags As Long
  dmDisplayFrequency As Long
End Type
```


### Source Code

```
'Changes the resolution to 640x480 with the current colordepth.
Dim DevM As DEVMODE
'Get the info into DevM
erg& = EnumDisplaySettings(0&, 0&, DevM)
'We don't change the colordepth, because a
'rebot will be necessary
DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT 'Or DM_BITSPERPEL
DevM.dmPelsWidth = 640 'ScreenWidth
DevM.dmPelsHeight = 480 'ScreenHeight
'DevM.dmBitsPerPel = 32 (could be 8, 16, 32 or even 4)
'Now change the display and check if possibleerg& = ChangeDisplaySettings(DevM, CDS_TEST)
'Check if succesfullSelect Case erg&
Case DISP_CHANGE_RESTART
  an = MsgBox("You've to reboot", vbYesNo + vbSystemModal, "Info")
  If an = vbYes Then
    erg& = ExitWindowsEx(EWX_REBOOT, 0&)
  End If
Case DISP_CHANGE_SUCCESSFUL
  erg& = ChangeDisplaySettings(DevM, CDS_UPDATEREGISTRY)
  MsgBox "Everything's ok", vbOKOnly + vbSystemModal, "It worked!"
Case Else
  MsgBox "Mode not supported", vbOKOnly + vbSystemModal, "Error"
End SelectEnd Sub
```

