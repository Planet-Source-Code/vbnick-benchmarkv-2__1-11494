VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MODEM"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4275
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   4275
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   660
      Top             =   4545
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   60
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton Command1 
      Caption         =   "go back"
      Height          =   480
      Left            =   165
      TabIndex        =   2
      Top             =   4590
      Width           =   3900
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   3930
      Left            =   60
      TabIndex        =   0
      Top             =   585
      Width           =   4125
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MODEM PROPERTIES"
      Height          =   450
      Left            =   30
      TabIndex        =   1
      Top             =   285
      Width           =   4050
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'we have used the sendMessage API here to add a horizontal scroll bar to the list...

Private Sub Command1_Click()
Unload Me
Form1.Show
End Sub

Private Sub form_load()
Show
'from here to-------------------------------------I
Dim x, uz                                        'I
uz = 2 * List1.Width / Screen.TwipsPerPixelX     'I
x = SendMessage(List1.hwnd, &H415, uz, 0)        'I
'here we added a horizontal scroll bar to list box.
On Local Error GoTo a
MSComm1.CommPort = 3
MSComm1.PortOpen = True
MSComm1.Output = "ATI4" & vbCr
MSComm1.InBufferCount = 0
MSComm1.InputLen = 0
Do
Loop While MSComm1.InBufferCount <= 0
List1.AddItem ("Modem Info:  " & MSComm1.Input)
MSComm1.Output = "ATI6" & vbCr
MSComm1.InBufferCount = 0
Do
Loop While MSComm1.InBufferCount <= 0
List1.AddItem ("Supported Standards: " & MSComm1.Input)
a:
MsgBox "modem port is already open", vbOKOnly, "TM"
'if you are connected to internet or something like that, this would handle the error message
If Winsock1.State <> 7 Then
Winsock1.Close

End If
End Sub
