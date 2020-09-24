VERSION 5.00
Begin VB.Form frmSets 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configure PingBall Challange"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4605
   Icon            =   "frmSets.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   4605
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstHardware 
      Height          =   2010
      Left            =   75
      TabIndex        =   4
      Top             =   630
      Width           =   4470
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save && Quit"
      Height          =   390
      Left            =   3060
      TabIndex        =   2
      Top             =   2790
      Width           =   1470
   End
   Begin VB.ComboBox cmbArray 
      Height          =   315
      Left            =   915
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   60
      Width           =   3630
   End
   Begin VB.Label Label2 
      Caption         =   "Misc:"
      Height          =   210
      Left            =   75
      TabIndex        =   3
      Top             =   420
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Resolution:"
      Height          =   210
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   825
   End
End
Attribute VB_Name = "frmSets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type DisplayModeDesc
    Width As Long
    Height As Long
    BPP As Byte
End Type
Private arr_DisplayModes() As DisplayModeDesc
Dim dx As New DirectX7
Dim binit As Boolean
Dim dd As DirectDraw7
Dim Mainsurf As DirectDrawSurface7
Dim primary As DirectDrawSurface7
Dim backbuffer As DirectDrawSurface7
Dim ddsd1 As DDSURFACEDESC2
Dim ddsd2 As DDSURFACEDESC2
Dim ddsd3 As DDSURFACEDESC2
Dim ddsd4 As DDSURFACEDESC2
Dim brunning As Boolean
Dim CurModeActiveStatus As Boolean
Dim bRestore As Boolean

Private Function GetAdapterInfo() As String

  Dim info As DirectDrawIdentifier

    Set dd = dx.DirectDrawCreate("")
    Set info = dd.GetDeviceIdentifier(DDGDI_DEFAULT)
    GetAdapterInfo = info.GetDescription
    Set dd = Nothing
End Function

Private Sub Command1_Click()
    If cmbArray.Text <> "" Then
        SaveSetting "Ping", "Res", "res", cmbArray.Text
    End If
    End
End Sub

Private Sub Form_Load()

    GetDisplayModes
    GetDDCaps

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    End

End Sub

Sub GetDisplayModes()

  'This is the actual code that reports back what display modes are available.
  'You could modify this to be a function, which enumerates the display modes
  'and runs through the list until it finds the one that you want. ie. You're program
  'runs in 800x600 in 32bpp mode; create a function that searches through the
  'available modes UNTIL it finds the one you want (800x600x32), at this point it
  'reports back True or false......

  Dim DisplayModesEnum As DirectDrawEnumModes
  Dim ddsd2 As DDSURFACEDESC2
  Dim dd As DirectDraw7 'These two lines can also be seen in the Init sub. This time

    'it doesn't go to fullscreen mode
    Set dd = dx.DirectDrawCreate("")
    dd.SetCooperativeLevel Me.hWnd, DDSCL_NORMAL

    'Create the Enumeration object
    Set DisplayModesEnum = dd.GetDisplayModesEnum(0, ddsd2)
    'Remember the array that wasn't defined? At this point
    'we set the size of the array.
    ReDim arr_DisplayModes(DisplayModesEnum.GetCount()) As DisplayModeDesc

    'This loop runs through the display modes, retrieving the data.
    'Height/Width/BPP aren't the only things that you can retrieve here....
    For i = 1 To DisplayModesEnum.GetCount()
        DisplayModesEnum.GetItem i, ddsd2
        If ddsd2.lWidth >= 640 And ddsd2.lHeight >= 480 And ddsd2.ddpfPixelFormat.lRGBBitCount > 8 Then
        cmbArray.AddItem CStr(ddsd2.lWidth) & "x" & CStr(ddsd2.lHeight) & "x" & ddsd2.ddpfPixelFormat.lRGBBitCount
        End If
        'cmbArray.Text = CStr(ddsd2.lWidth) & "x" & CStr(ddsd2.lHeight) & " " & CStr(ddsd2.ddpfPixelFormat.lRGBBitCount) & "bpp"

        'This fills out the data structure to include information
        'on the current display mode.
        arr_DisplayModes(i).Height = ddsd2.lHeight
        arr_DisplayModes(i).Width = ddsd2.lWidth
        arr_DisplayModes(i).BPP = ddsd2.ddpfPixelFormat.lRGBBitCount
    Next i

    'Directdraw is no longer needed - destroy it.
    Set dd = Nothing

End Sub

Sub GetDDCaps()

  'This part returns the capabilities of DirectDraw
  'You only really need this information if you're going to do
  'any technical stuff.

  Dim dd As DirectDraw7
  Dim hwCaps As DDCAPS 'HARDWARE
  Dim helCaps As DDCAPS 'SOFTWARE EMULATION

    Set dd = dx.DirectDrawCreate("")
    dd.SetCooperativeLevel Me.hWnd, DDSCL_NORMAL
    dd.GetCaps hwCaps, helCaps

    'how much video memory is available
    lstHardware.AddItem "GENERAL INFORMATION"
    lstHardware.AddItem GetAdapterInfo
    'The memory amount can be useful. If you know that you're surfaces require
    '450kb of memory then you can check if the host computer has this much memory.
    lstHardware.AddItem " total video memory " & CStr(hwCaps.lVidMemTotal) & " bytes (" & CStr(Format$(hwCaps.lVidMemTotal / 1024, "#.0")) & "Kb)"
    lstHardware.AddItem " free video memory " & CStr(hwCaps.lVidMemFree) & " bytes (" & CStr(Format$(hwCaps.lVidMemFree / 1024, "#.0")) & "Kb)"

    lstHardware.AddItem " There are " & hwCaps.lNumFourCCCodes & " FourCC codes available"

    lstHardware.AddItem ""

    lstHardware.AddItem "HARDWARE CAPABILITIES"

    'You can get a list of what these constants mean in the
    'sdk help file. If you don't have the help file you're a bit stuck!

    lVal = hwCaps.ddsCaps.lCaps2
    If lVal And DDCAPS2_CANCALIBRATEGAMMA Then
        lstHardware.AddItem " Supports gamma correction"
      Else
        lstHardware.AddItem " No support for gamma correction"
    End If

    If lVal And DDCAPS2_CERTIFIED Then
        lstHardware.AddItem "The driver is certified"
      Else
        lstHardware.AddItem " The driver is not certified"
    End If

    If lVal And DDCAPS2_WIDESURFACES Then
        lstHardware.AddItem " support for surfaces wider than the screen"
      Else
        lstHardware.AddItem " No support for surfaces wider than the screen"
    End If

    lVal = hwCaps.lSVBFXCaps
    If lVal And DDFXCAPS_BLTALPHA Then
        lstHardware.AddItem " Support for Alpha Blended Blit operations"
      Else
        lstHardware.AddItem " No support for Alpha Blended Blit operations"
    End If

    If lVal And DDFXCAPS_BLTROTATION Then
        lstHardware.AddItem " Support for rotation Blit operations"
      Else
        lstHardware.AddItem " No support for rotation Blit operations"
    End If

    lVal = hwCaps.lSSBCaps
    If lVal And DDCAPS_3D Then
        lstHardware.AddItem " Support for 3D Acceleration"
      Else
        lstHardware.AddItem " No support for 3D acceleration"
    End If

    If lVal And DDCAPS_BLTQUEUE Then
        lstHardware.AddItem " Support for asynchronous blitting"
      Else
        lstHardware.AddItem " No support for asynchronous blitting"
    End If

    If lVal And DDCAPS_BLTSTRETCH Then
        lstHardware.AddItem " Support for stretching during Blit operations"
      Else
        lstHardware.AddItem " No support for stretching during blit operations"
    End If

    If lVal And DDCAPS_NOHARDWARE Then
        lstHardware.AddItem " Hardware support is available"
      Else
        lstHardware.AddItem " No hardware support"
    End If

    '//////////////////////SOFTWARE\\\\\\\\\\\\\\\\\\\\\\\\\\\\

    lstHardware.AddItem "SOFTWARE CAPABILITIES"

    lVal = helCaps.ddsCaps.lCaps2
    If lVal And DDCAPS2_WIDESURFACES Then
        lstHardware.AddItem " The device supports surfaces wider than the screen"
      Else
        lstHardware.AddItem " The device does not support surfaces wider than the screen"
    End If

    lVal = helCaps.lSVBFXCaps
    If lVal And DDFXCAPS_BLTALPHA Then
        lstHardware.AddItem " Software supports Alpha Blended Blit operations"
      Else
        lstHardware.AddItem " No Software support for Alpha Blended Blit operations"
    End If

    If lVal And DDFXCAPS_BLTROTATION Then
        lstHardware.AddItem " Software supports rotation Blit operations"
      Else
        lstHardware.AddItem " No software support for rotation Blit operations"
    End If

    lVal = helCaps.lSSBCaps
    If lVal And DDCAPS_3D Then
        lstHardware.AddItem " Software supports 3D Acceleration"
      Else
        lstHardware.AddItem " No software support for 3D acceleration"
    End If

    If lVal And DDCAPS_BLTQUEUE Then
        lstHardware.AddItem " Software supports asynchronous blitting"
      Else
        lstHardware.AddItem " No software support for asynchronous blitting"
    End If

    Set dd = Nothing

End Sub

':) VB Code Formatter V2.12.7 (23/06/2002 17:34:49) 20 + 195 = 215 Lines
