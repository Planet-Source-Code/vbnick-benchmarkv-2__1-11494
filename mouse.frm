VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MOUSE PROPERTIES"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3360
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   3360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2505
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   1305
      Width           =   825
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CHANGE DOUBLE CLICKING TIME"
      Height          =   465
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   3345
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Left            =   2535
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   465
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CHANGE CURSOR BLINK RATE"
      Height          =   420
      Left            =   15
      TabIndex        =   0
      Top             =   60
      Width           =   3300
   End
   Begin VB.Label Label2 
      Caption         =   "Value"
      Height          =   360
      Left            =   1665
      TabIndex        =   5
      Top             =   1395
      Width           =   765
   End
   Begin VB.Label Label1 
      Caption         =   "Value"
      Height          =   300
      Left            =   1635
      TabIndex        =   4
      Top             =   540
      Width           =   1050
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Shell "rundll32 user,setcaretblinktime val(text1)"
Text1.SetFocus
End Sub

Private Sub Command2_Click()
Shell "rundll32 user,setdoubleclicktime val(text2)"
Text2.SetFocus
End Sub

Private Sub Form_Load()
Text1.Text = 10
Text2.Text = 10
End Sub
