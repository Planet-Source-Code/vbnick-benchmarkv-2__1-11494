VERSION 5.00
Begin VB.Form klavye 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Keyboard Information"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   Icon            =   "klavye.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   4650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Go Back"
      Height          =   495
      Left            =   60
      TabIndex        =   4
      Top             =   3420
      Width           =   4560
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Re-Test"
      Height          =   495
      Left            =   75
      TabIndex        =   3
      Top             =   2895
      Width           =   1860
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2130
      Left            =   30
      TabIndex        =   0
      Top             =   705
      Width           =   4590
   End
   Begin VB.Label Label2 
      Caption         =   "Keyboard Information"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   90
      TabIndex        =   2
      Top             =   210
      Width           =   3600
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Press any button or just click."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   480
      Left            =   1980
      TabIndex        =   1
      Top             =   2895
      Width           =   2640
   End
   Begin VB.Shape Shape1 
      Height          =   630
      Left            =   3780
      Shape           =   4  'Rounded Rectangle
      Top             =   30
      Width           =   825
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3975
      Picture         =   "klavye.frx":0442
      Top             =   105
      Width           =   480
   End
End
Attribute VB_Name = "klavye"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetKBCodePage Lib "user32" () As Long
Private Declare Function GetGraphicsMode Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long

Private Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long

Private Declare Function GetKeyboardLayoutName Lib "user32" Alias "GetKeyboardLayoutNameA" (ByVal pwszKLID As String) As Long

Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long

Private Declare Function GetKeyboardType Lib "user32" (ByVal nTypeFlag As Long) As Long

Private Declare Function GetKeyNameText Lib "user32" Alias "GetKeyNameTextA" (ByVal lParam As Long, ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetInputState Lib "user32" () As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Sub Command2_Click()
Unload Me
Load Form1
Form1.Show
End Sub

Private Sub form_load()
Show
List1.AddItem ("                   KEYBOARD INFORMATION")
List1.AddItem ("The Code Page is " & GetKBCodePage() & "")
Dim a, str1, b, c
a = GetKeyboardType(0)
b = GetKeyboardLayout(0)
Select Case a
Case 1: str1 = "XT type 83 button "
Case 2: str1 = "Olivetti type 102 button "
Case 3: str1 = "AT type 84 button "
Case 4: str1 = "101 or 102 button "
Case 5: str1 = "Nokia 1050 type "
Case 6: str1 = "Nokia 9140 type "
Case Else: str1 = "Unknown  Type "
End Select
List1.AddItem (str1 & Str(GetKeyboardType(2)) & " function button keyboard")

Timer1.Interval = 10
End Sub

Private Sub Timer1_Timer()
Dim i
Label1.Caption = "started"
While GetInputState() = 0
Wend
Label1.Caption = "The keyboard is tested sucessfully, no errors were found."

End Sub
