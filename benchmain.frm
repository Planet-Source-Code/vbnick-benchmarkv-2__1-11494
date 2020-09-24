VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BENCHMARK PROGRAM V 2.12 BY KAYHAN TANRISEVEN"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
   Icon            =   "benchmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command14 
      Caption         =   "OPEN"
      Height          =   330
      Left            =   3525
      TabIndex        =   23
      Top             =   5040
      Width           =   1470
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "benchmain.frx":0442
      Left            =   4980
      List            =   "benchmain.frx":0458
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   5040
      Width           =   1710
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Open Control Panel"
      Height          =   420
      Left            =   75
      TabIndex        =   21
      Top             =   5025
      Width           =   1875
   End
   Begin VB.Timer Timer2 
      Left            =   6120
      Top             =   3720
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Programs"
      Height          =   840
      Left            =   120
      Picture         =   "benchmain.frx":048B
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3405
      Width           =   3555
   End
   Begin VB.CommandButton Command11 
      Height          =   525
      Left            =   6150
      Picture         =   "benchmain.frx":08CD
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   -15
      Width           =   555
   End
   Begin VB.Timer Timer1 
      Left            =   6255
      Top             =   5025
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2310
      Left            =   3810
      TabIndex        =   12
      Top             =   525
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   4075
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   0
      MousePointer    =   1
      ShowWeekNumbers =   -1  'True
      StartOfWeek     =   24576001
      CurrentDate     =   36754
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   30
      ScaleHeight     =   450
      ScaleWidth      =   4965
      TabIndex        =   1
      Top             =   25
      Width           =   4965
      Begin VB.CommandButton Command10 
         Height          =   450
         Left            =   4455
         Picture         =   "benchmain.frx":0D0F
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   0
         Width           =   480
      End
      Begin VB.CommandButton Command9 
         Height          =   450
         Left            =   3960
         Picture         =   "benchmain.frx":0E99
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         Width           =   480
      End
      Begin VB.CommandButton Command8 
         Height          =   450
         Left            =   3435
         Picture         =   "benchmain.frx":11A3
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         Width           =   510
      End
      Begin VB.CommandButton Command7 
         Height          =   450
         Left            =   2940
         Picture         =   "benchmain.frx":15E5
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         Width           =   480
      End
      Begin VB.CommandButton Command6 
         Height          =   450
         Left            =   2445
         Picture         =   "benchmain.frx":1A27
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   480
      End
      Begin VB.CommandButton Command5 
         Height          =   450
         Left            =   1950
         Picture         =   "benchmain.frx":1E69
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   480
      End
      Begin VB.CommandButton Command4 
         Height          =   450
         Left            =   1470
         Picture         =   "benchmain.frx":22AB
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   480
      End
      Begin VB.CommandButton Command3 
         Height          =   450
         Left            =   975
         Picture         =   "benchmain.frx":26ED
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   480
      End
      Begin VB.CommandButton Command2 
         Height          =   450
         Left            =   480
         Picture         =   "benchmain.frx":2B2F
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   480
      End
      Begin VB.CommandButton Command1 
         Height          =   450
         Left            =   0
         Picture         =   "benchmain.frx":2F71
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   480
      End
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   510
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6720
      _ExtentX        =   11853
      _ExtentY        =   900
      BandCount       =   2
      Picture         =   "benchmain.frx":33B3
      BackColor       =   16777215
      EmbossPicture   =   -1  'True
      _CBWidth        =   6720
      _CBHeight       =   510
      _Version        =   "6.0.8169"
      MinHeight1      =   450
      Width1          =   6630
      NewRow1         =   0   'False
      BandStyle1      =   1
      MinHeight2      =   360
      NewRow2         =   0   'False
      BandStyle2      =   1
      Begin VB.CommandButton Command15 
         Caption         =   "Perfonmance Graphics"
         Height          =   480
         Left            =   4965
         TabIndex        =   24
         Top             =   15
         Width           =   1200
      End
   End
   Begin VB.Label Label7 
      Caption         =   "kayahan@presidency.com"
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   960
      TabIndex        =   20
      Top             =   2580
      Width           =   2070
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Height          =   600
      Left            =   105
      Top             =   4380
      Width           =   6570
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00404040&
      BorderStyle     =   6  'Inside Solid
      Height          =   660
      Left            =   60
      Top             =   4350
      Width           =   6630
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "COMPUTER WORKING HOUR COUNTER"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1830
      TabIndex        =   19
      Top             =   4425
      Width           =   3480
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   135
      TabIndex        =   18
      Top             =   4695
      Width           =   6465
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C0C0C0&
      X1              =   30
      X2              =   6735
      Y1              =   4305
      Y2              =   4305
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00404040&
      X1              =   30
      X2              =   6705
      Y1              =   4275
      Y2              =   4275
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H8000000C&
      Height          =   1335
      Left            =   3810
      Shape           =   4  'Rounded Rectangle
      Top             =   2910
      Width           =   2895
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H8000000C&
      Height          =   2745
      Left            =   60
      Shape           =   4  'Rounded Rectangle
      Top             =   615
      Width           =   3495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"benchmain.frx":B425
      Height          =   2595
      Left            =   120
      TabIndex        =   16
      Top             =   705
      Width           =   3360
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "For time and date settings please use the button above , with a little clock picture on it."
      Height          =   1080
      Left            =   4020
      TabIndex        =   15
      Top             =   3210
      Width           =   2625
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000000&
      Height          =   435
      Left            =   30
      Shape           =   4  'Rounded Rectangle
      Top             =   5040
      Width           =   6630
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000000&
      Index           =   1
      X1              =   3735
      X2              =   3735
      Y1              =   195
      Y2              =   4275
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      Index           =   0
      X1              =   3705
      X2              =   3705
      Y1              =   195
      Y2              =   4275
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3975
      TabIndex        =   13
      Top             =   2910
      Width           =   2700
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetTickCount& Lib "kernel32" ()
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long

Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Integer, ByVal bRevert As Integer) As Integer
Dim sh, x, tanitici1, tanitici2, tanitici3

Private Sub Command1_Click()
Load driver
driver.Show

End Sub

Private Sub Command10_Click()
Load sounds
sounds.Show

End Sub

Private Sub Command11_Click()
MsgBox "will be ready soon!!", vbOKOnly, "TM"
End Sub

Private Sub Command12_Click()
Load Form2
Form2.Show

End Sub

Private Sub Command13_Click()
Shell "rundll32 shell32,ControlP_RunDLL", vbNormalFocus

End Sub

Private Sub Command14_Click()
Shell "rundll32 shell32,Control_RunDLL main.cpl @" & Combo1.ItemData(Combo1.ListIndex)

End Sub

Private Sub Command15_Click()
Load grafik
grafik.Show
End Sub

Private Sub Command2_Click()
driver.Show
End Sub

Private Sub Command3_Click()
driver.Show
End Sub

Private Sub command4_click()
driver.Show
End Sub

Private Sub Command5_Click()
Load klavye
klavye.Show

End Sub

Private Sub Command6_Click()
Load Form4
Form4.Show

End Sub

Private Sub Command7_Click()
Load EKRAN
EKRAN.Show
End Sub

Private Sub Command8_Click()
Load printer
printer.Show
End Sub

Private Sub Command9_Click()
Load Form3
Form3.Show
End Sub

Private Sub Form_Load()
'BY THE WAY CHECK OUT THE ICON ON THE LEFT, TITLE BAR....
'I HAVE ADDED SOME INFO TO THAT SYSTEM MENU....
'AND YOU CAN SEE HOW ITS DONE HERE..

Timer1.Interval = 100
Timer2.Interval = 200
tanitici1 = 10
tanitici2 = 20
tanitici3 = 30
sh = GetSystemMenu(Form1.hwnd, False)
x = AppendMenu(sh, 0, tanitici1, "     This program has programmed")
x = AppendMenu(sh, 0, tanitici2, "       by Kayhan Tanriseven")
x = AppendMenu(sh, 0, tanitici3, "        PLEASE VOTE FOR ME.")
Combo1.ListIndex = 0
End Sub
Private Sub Timer1_Timer()
Label1.Caption = "Time: " & Time
End Sub

Private Sub Timer2_Timer()
'THIS TIME WE WILL GET THE TICKCOUNT OF THE COMPUTER...
'THEN WILL DVIDE THEM IN ORDER TO GET THE FORMATTED TIME..

Dim t
t = GetTickCount&
Label5.Caption = "Since " & (t \ 1440000) & ":"
t = t Mod (60000 * 24)
Label5.Caption = Label5.Caption & (t \ 60000) & ":"
t = t Mod (60000)
Label5.Caption = Label5.Caption & (t \ 1000) & " 'the computer is working"


End Sub
