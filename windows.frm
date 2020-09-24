VERSION 5.00
Begin VB.Form EKRAN 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Screen Information"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3120
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   162
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "windows.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   3120
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "REPAINT SCREEN"
      Height          =   405
      Left            =   105
      TabIndex        =   4
      Top             =   3510
      Width           =   2970
   End
   Begin VB.CommandButton Command2 
      Caption         =   "DISPLAY PROPERTIES"
      Height          =   435
      Left            =   105
      TabIndex        =   3
      Top             =   3915
      Width           =   2970
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go Back"
      Height          =   435
      Left            =   105
      TabIndex        =   1
      Top             =   4350
      Width           =   2970
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   2970
      Left            =   60
      TabIndex        =   0
      Top             =   420
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "SCREEN INFO"
      Height          =   345
      Left            =   -45
      TabIndex        =   2
      Top             =   165
      Width           =   2970
   End
End
Attribute VB_Name = "EKRAN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Sub Command2_Click()
Shell "rundll32 shell32,Control_RunDLL desk.cpl"
End Sub

Private Sub Command3_Click()
Shell "rundll32 user,repaintscreen"
End Sub

Private Sub form_load()
Show
Dim x As String, ver, tek
FontUnderline = False
x = String(0, 255)
ver = GetDeviceCaps(hdc, 0)

List1.AddItem ("")
List1.AddItem ("Version:  " & (ver And &HFF00) / &H100) & "." & (ver And &HFF)
tek = GetDeviceCaps(hdc, 2)
Select Case tek
Case 0: List1.AddItem ("Teknology:    Vektoral Plotter")
Case 1: List1.AddItem ("Teknology:    Raster Display")
Case 2: List1.AddItem ("Teknology:    Raster Printer")
Case 3: List1.AddItem ("Teknology:    Raster Camera")
Case 4: List1.AddItem ("Teknology:    Character Stream, PLP")
Case 5: List1.AddItem ("Teknology:    Metafile, VDM")
Case 6: List1.AddItem ("Teknology:    Display File")
End Select
List1.AddItem ("Horizontal:  " & GetDeviceCaps(hdc, 4) & " mm")
List1.AddItem ("Vertical:  " & GetDeviceCaps(hdc, 6) & " mm")
List1.AddItem ("Resolution:  " & GetDeviceCaps(hdc, 8) & " X " & GetDeviceCaps(hdc, 10))
'List1.AddItem ("Color Count:  " & GetDeviceCaps(hdc, 14) * 2 ^ GetDeviceCaps(hdc, 12))
Dim g
g = GetDeviceCaps(hdc, 14) * 2 ^ GetDeviceCaps(hdc, 12)
Select Case g
Case "4294967296": List1.AddItem ("Color Count:  32 Bit")
Case "16777216": List1.AddItem ("Color Count:  24 Bit")
Case "65536": List1.AddItem ("Color Count:  16 Bit")
Case "256": List1.AddItem ("Color Count:  256 Color")
End Select

Dim d As String, t As Integer, s

d = String(255, 0)
't = GetPrivateProfileString("windows", "device", "", d, Len(d), "win.ini")
t = GetPrivateProfileString("boot.description", "display.drv", "", d, Len(d), "system.ini")

If t = 0 Then
s = "unknown"
Else
s = Left(d, t)
End If
List1.AddItem ("Graphics Card:    " & s)
Dim w, uz
uz = 2 * List1.Width / Screen.TwipsPerPixelX
w = SendMessage(List1.hwnd, &H415, uz, 0)


Dim e As String, f As Integer, r

e = String(255, 0)
 
f = GetPrivateProfileString("desktop", "wallpaper", "", e, Len(e), "win.ini")

If f = 0 Then
r = "unknown"
Else
r = Left(e, f)
End If
List1.AddItem ("Wallpaper:    " & r)
End Sub
