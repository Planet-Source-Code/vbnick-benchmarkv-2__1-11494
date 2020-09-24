VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form sounds 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "THE SOUND SYSTEM INFORMATION"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6360
   Icon            =   "sounds.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Sound Recorder Information"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2205
      Left            =   3450
      TabIndex        =   7
      Top             =   2370
      Width           =   2835
      Begin VB.CommandButton Command2 
         Caption         =   "Microphone Settings"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1515
         TabIndex        =   10
         Top             =   1650
         Width           =   1260
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Sound Recorder "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   45
         TabIndex        =   9
         Top             =   1650
         Width           =   1350
      End
      Begin VB.Timer Timer1 
         Left            =   300
         Top             =   1200
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1725
         Left            =   105
         TabIndex        =   8
         Top             =   360
         Width           =   2610
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Midi Information"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   3465
      TabIndex        =   6
      Top             =   75
      Width           =   2850
      Begin VB.CommandButton Command3 
         Caption         =   "Midi Settings"
         Height          =   450
         Left            =   90
         TabIndex        =   13
         Top             =   1725
         Width           =   2745
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1110
         Left            =   120
         TabIndex        =   11
         Top             =   315
         Width           =   2595
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Genreral Info"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4500
      Left            =   90
      TabIndex        =   5
      Top             =   75
      Width           =   3300
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3810
         Left            =   60
         TabIndex        =   12
         ToolTipText     =   "Ses Sisteminiz"
         Top             =   225
         Width           =   3165
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Left && Right Speaker  Settings"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   4605
      Width           =   6150
      Begin MSComctlLib.Slider Slider1 
         Height          =   495
         Left            =   720
         TabIndex        =   1
         Top             =   255
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   873
         _Version        =   393216
      End
      Begin MSComctlLib.Slider Slider2 
         Height          =   495
         Left            =   3660
         TabIndex        =   2
         Top             =   240
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   873
         _Version        =   393216
      End
      Begin VB.Label Label2 
         Caption         =   "Left"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3030
         TabIndex        =   4
         Top             =   330
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "Right"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   60
         TabIndex        =   3
         Top             =   330
         Width           =   720
      End
   End
End
Attribute VB_Name = "sounds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAXPNAMELEN = 32  '  en fazla urun adý uzunlugu 32 dýr

Private Type WAVEOUTCAPS
        wMid As Integer
        wPid As Integer
        vDriverVersion As Long
        szPname As String * MAXPNAMELEN
        dwFormats As Long
        wChannels As Integer
        dwSupport As Long
End Type

Private Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long

Private Declare Function waveOutGetDevCaps Lib "winmm.dll" Alias "waveOutGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As WAVEOUTCAPS, ByVal uSize As Long) As Long

Private Declare Function waveOutGetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, lpdwVolume As Long) As Long
Private Declare Function waveOutSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Private Declare Function waveInGetNumDevs Lib "winmm.dll" () As Long
Private Declare Function midiInGetNumDevs Lib "winmm.dll" () As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Sub form_load()
Dim lpc As WAVEOUTCAPS
Dim uz, d
uz = 2 * List1.Width / Screen.TwipsPerPixelX
d = SendMessage(List1.hwnd, &H415, uz, 0)

If waveOutGetNumDevs() = 0 Then
Frame1.Caption = "Right && Left Speaker Settings- The drive is not present or supported."
Frame2.Caption = "General Information - Sound Card is not dedected"
Exit Sub
End If
List1.AddItem ("In your System, You have " & waveOutGetNumDevs() & " sound card installed")

Call waveOutGetDevCaps(0, lpc, Len(lpc))
List1.AddItem ("S.Card Name :  " & lpc.szPname)

If lpc.wChannels = 1 Then
List1.AddItem ("Version : " & (lpc.vDriverVersion And &HFF00) \ 256 & "." & (lpc.vDriverVersion And &HFF) & "   Mono")
Else
List1.AddItem ("Version : " & (lpc.vDriverVersion And &HFF00) \ 256 & "." & (lpc.vDriverVersion And &HFF) & "   Stereo")
End If
List1.AddItem ("********************************************")
List1.AddItem ("            Supported Formats")
List1.AddItem ("********************************************")
If lpc.dwFormats And 1 Then List1.AddItem (" 11.025 kHz ,Mono, 8-Bit")
If lpc.dwFormats And 2 Then List1.AddItem (" 11.025 kHz ,Stereo, 8-Bit")
If lpc.dwFormats And 4 Then List1.AddItem (" 11.025 kHz ,Mono, 16-Bit")
If lpc.dwFormats And 8 Then List1.AddItem (" 11.025 kHz ,Stereo, 8-Bit")
If lpc.dwFormats And &H10 Then List1.AddItem (" 22.05 kHz ,Mono, 8-Bit")
If lpc.dwFormats And &H20 Then List1.AddItem (" 22.05 kHz ,Stereo, 8-Bit")
If lpc.dwFormats And &H40 Then List1.AddItem (" 22.05 kHz ,Mono, 16-Bit")
If lpc.dwFormats And &H80 Then List1.AddItem (" 22.05 kHz ,Stereo, 16-Bit")
If lpc.dwFormats And &H100 Then List1.AddItem (" 44.01 kHz ,Mono, 8-Bit")
If lpc.dwFormats And &H200 Then List1.AddItem (" 22.05 kHz ,Stereo, 8-Bit")
If lpc.dwFormats And &H400 Then List1.AddItem (" 22.05 kHz ,Mono, 16-Bit")
If lpc.dwFormats And &H800 Then List1.AddItem (" 22.05 kHz ,Stereo, 16-Bit")
List1.AddItem ("----------------------------------------------")
If lpc.dwSupport And 1 Then List1.AddItem ("Pitch Level Controller is Installed") Else List1.AddItem ("Pitch Level Controller is not Installed")
If lpc.dwSupport And 4 Then List1.AddItem ("Independent Speaker Control is installed") Else List1.AddItem (" Independent Speaker Control is not istalled")
If lpc.dwSupport And 8 Then List1.AddItem ("Right - Left Speaker Sound control is supported") Else List1.AddItem ("Right - Left Speaker Sound Control is not Supported")
If lpc.dwSupport And 2 Then List1.AddItem ("Playback rate Control is supported") Else List1.AddItem ("Playback rate Control is not supported")

If lpc.wChannels = 0 Then
Slider2.Enabled = False
Frame1.Caption = "Right && Left Sound Control - The device is not supported."
End If
If (lpc.dwSupport And 4) = 0 Then
Slider1.Enabled = False
Slider2.Enabled = False
Frame1.Caption = "Right && Left Sound Control - The device is not supported."
End If
If (lpc.dwSupport And 8) = 0 Then
Slider2.Enabled = False
End If
Slider1.Min = 0
Slider1.Max = &HFFFF&
Slider1.TickFrequency = &HFFFF& / 10
Slider2.Min = 0
Slider2.Max = &HFFFF&
Slider2.TickFrequency = &HFFFF& / 10
Dim x, sol, sag, st
Call waveOutGetVolume(0, x)
sol = x And &HFFFF&
st = Hex(x And &HFFFF0000)
If Len(st) > 4 Then
st = Mid(st, 1, Len(st) - 4)
Else
st = "0"
End If
sag = CDbl("&h" & st)
Slider1.Value = sol
Slider2.Value = sag
Timer1.Enabled = True
Timer1.Interval = 100
If midiInGetNumDevs = 0 Then
Label4.Caption = "In your system, Midi && Midi Controls is not supported or Malufunctioning."
Else
Label4.Caption = "In your system, Midi && Midi Controls is  supported and working normally."
End If
End Sub

Sub sesayar()
Dim x, sol, sag, s
sol = Slider1.Value
sag = Slider2.Value
s = Val("&h" & Hex(sag) & String(4 - Len(Hex(sol)), "0") & Hex(sol) & "&")
Call waveOutSetVolume(0, s)
End Sub
Private Sub slider1_click()
sesayar
End Sub

Private Sub slider2_click()
sesayar
End Sub
Private Sub slider1_scroll()
sesayar
End Sub
Private Sub slider2_scroll()
sesayar
End Sub

Private Sub Timer1_Timer()
If waveInGetNumDevs() = 0 Then
Label3.Caption = "In your system, The Sound Recording function is not supported or malufunctioning."
Command1.Enabled = False
Command2.Enabled = False

Else
Label3.Caption = "In your system, The Sound Recording function is supported and working normally."
Command1.Enabled = True
Command2.Enabled = True

End If
Timer1.Enabled = False
End Sub
