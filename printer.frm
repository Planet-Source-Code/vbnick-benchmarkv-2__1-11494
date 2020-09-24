VERSION 5.00
Begin VB.Form Printer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Printer Information"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3345
   Icon            =   "printer.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   3345
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "PRINT TEST PAGE"
      Height          =   435
      Left            =   45
      TabIndex        =   3
      Top             =   3450
      Width           =   3285
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go Back"
      Height          =   435
      Left            =   45
      TabIndex        =   1
      Top             =   3885
      Width           =   3285
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   2955
      Left            =   45
      TabIndex        =   0
      Top             =   480
      Width           =   3270
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Printer Properties"
      Height          =   375
      Left            =   270
      TabIndex        =   2
      Top             =   120
      Width           =   2640
   End
End
Attribute VB_Name = "printer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Sub Command1_Click()
Load Form1
Form1.Show
Unload Me
End Sub

Private Sub Command2_Click()
Shell "rundll32 msprint2.dll,RUNDLL_PrintTestPage"
End Sub

Private Sub Form_Load()
Show
Dim x As String, ver, tek
FontUnderline = True
List1.AddItem ("         Printer Info")
FontUnderline = False
Dim d As String, t As Integer, s, m As Integer

d = String(255, 0)
t = GetPrivateProfileString("windows", "device", "", d, Len(d), "win.ini")
'm = GetPrivateProfileString("[boot.description", "keyboard.drv", x, Len(x), "system.ini")

If t = 0 Then
s = "unknown"
Else
s = Left(x, t)
End If
List1.AddItem ("Printer Name:    " & s)


x = String(255, 0)
ver = GetDeviceCaps(printer.hdc, 0)
List1.AddItem ("Version:   " & ((ver And &HFF00) / &H100) & "." & (ver And &HFF))
tek = GetDeviceCaps(printer.hdc, 2)
Select Case tek
Case 0: List1.AddItem ("Technology:    Vektoral Plotter")
Case 1: List1.AddItem ("Technology:    Raster Display")
Case 2: List1.AddItem ("Technology:    Raster Printer")
Case 3: List1.AddItem ("Technology:    Raster Camera")
Case 4: List1.AddItem ("Technology:    Character Stream, PLP")
Case 5: List1.AddItem ("Technology:    Metafile, VDM")
Case 6: List1.AddItem ("Technology:    Display File")
End Select

List1.AddItem ("Horizontal:  " & GetDeviceCaps(hdc, 4) & " mm")
List1.AddItem ("Vertical:  " & GetDeviceCaps(hdc, 6) & " mm")
List1.AddItem ("Resolution:  " & GetDeviceCaps(hdc, 8) & " X " & GetDeviceCaps(hdc, 10) & " Pixel")
'List1.AddItem ("Color Count:  " & g)
List1.AddItem ("Font Count:  " & GetDeviceCaps(hdc, 22))


'THIS PORTION OF THIS CODE WILL WORK ON SOME PRINTER TYPES, AND WONT WORK
'ON OTHERS..SO I COMMENTED THEM OUT BELOW BUT DIDNT LET THEM WORK, BECAUSE YOU
'CAN GET  AN INTERNAL ERROR MESSAGE IF YOUR PRINTER DOESNT SUPPORT FOLLOWING..
'=======================================================
'List1.AddItem ("Paper Bin: " & printer.PaperBin)      =
'List1.AddItem ("Color Mode: " & printer.ColorMode)    =
'List1.AddItem ("Copies: " & printer.Copies)           =
'List1.AddItem ("Device Name: " & printer.DeviceName)  =
'List1.AddItem ("Driver Name: " & printer.DriverName)  =
'List1.AddItem ("Duplex: " & printer.Duplex)           =
'List1.AddItem ("Orientation: " & printer.Orientation) =
'List1.AddItem ("Paper Size: " & printer.PaperSize)    =
'List1.AddItem ("Printer Port: " & printer.Port)       =
'List1.AddItem ("Print Quality: " & printer.Printqality)
'List1.AddItem ("Track Default: " & printer.TrackDefault)
'List1.AddItem ("Zoom: " & printer.Zoom)               =
'=======================================================


End Sub
