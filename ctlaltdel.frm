VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Task Manager"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   Icon            =   "ctlaltdel.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "ctlaltdel.frx":0442
   MousePointer    =   99  'Custom
   ScaleHeight     =   5640
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Height          =   4710
      Left            =   60
      TabIndex        =   1
      Top             =   45
      Visible         =   0   'False
      Width           =   5745
   End
   Begin VB.Timer Timer2 
      Left            =   3345
      Top             =   60
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "UPDATE"
      Height          =   420
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5175
      Width           =   5805
   End
   Begin VB.Timer Timer1 
      Left            =   2775
      Top             =   60
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long

Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long

Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal aint As Integer) As Integer

Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

Private Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long

Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Integer, ByVal bRevert As Integer) As Integer
Dim sh, x, tanitici1, tanitici2, tanitici3
Private Sub Command1_Click()
Form2.Cls
Timer1.Enabled = True
Timer1.Interval = 50
End Sub

Private Sub form_load()
Timer1.Interval = 500
Timer2.Interval = 200
tanitici1 = 10
tanitici2 = 20
tanitici3 = 30
sh = GetSystemMenu(Form2.hwnd, False)
x = AppendMenu(sh, 0, tanitici1, "     Bu Program Kayhan")
x = AppendMenu(sh, 0, tanitici2, "   Tanriseven Tarafindan")
x = AppendMenu(sh, 0, tanitici3, "       Programlanmistir.")
End Sub

Private Sub Timer1_Timer()
Show
Dim Aktif_pencere_hwnd, uz As Integer, baslik As String
Aktif_pencere_hwnd = GetWindow(Form2.hwnd, 0)
While Aktif_pencere_hwnd <> 0

uz = GetWindowTextLength(Aktif_pencere_hwnd)
baslik = Space(uz + 1)
uz = GetWindowText(Aktif_pencere_hwnd, baslik, uz + 1)
If uz > 0 Then
baslik = Left(baslik, uz)
If IsIconic(Aktif_pencere_hwnd) Then
Print baslik & " -->minimized"
List1.AddItem baslik & " -->minimized"
End If
If IsZoomed(Aktif_pencere_hwnd) Then
Print baslik & " -->maximized"
List1.AddItem baslik & " -->maximized"
End If
If IsWindowVisible(Aktif_pencere_hwnd) Then
Print baslik & " --> not active"
List1.AddItem baslik & " -->not active"

End If
End If
Aktif_pencere_hwnd = GetWindow(Aktif_pencere_hwnd, 2)

Wend
Timer1.Enabled = False
End Sub

Private Sub Form_DblClick()
On Local Error Resume Next
AppActivate List1.Text
SendKeys "% "
SendKeys "{ENTER}"

End Sub
Private Sub form_unload(cancel As Integer)
Load Form1
Form1.Visible = True
Form1.Show

End Sub

Private Sub Timer2_Timer()
Call FlashWindow(Form2.hwnd, True)
End Sub
