VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form grafik 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grafical Perfonmance Calculator By Kayhan Tanriseven"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8775
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   162
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "grafik.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   8775
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Calculate"
      Height          =   450
      Left            =   4785
      TabIndex        =   12
      Top             =   5580
      Width           =   3420
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Calculate"
      Height          =   465
      Left            =   180
      TabIndex        =   11
      Top             =   5595
      Width           =   3720
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Calculate"
      Height          =   420
      Left            =   4800
      TabIndex        =   10
      Top             =   2595
      Width           =   3510
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate"
      Default         =   -1  'True
      Height          =   450
      Left            =   255
      TabIndex        =   9
      Top             =   2595
      Width           =   3585
   End
   Begin VB.DriveListBox Drive1 
      Height          =   330
      Left            =   5790
      TabIndex        =   4
      Top             =   3330
      Width           =   2010
   End
   Begin MSChart20Lib.MSChart MSChart4 
      Height          =   2385
      Left            =   4485
      OleObjectBlob   =   "grafik.frx":0442
      TabIndex        =   3
      Top             =   3435
      Width           =   3930
   End
   Begin MSChart20Lib.MSChart MSChart3 
      Height          =   2400
      Left            =   135
      OleObjectBlob   =   "grafik.frx":279A
      TabIndex        =   2
      Top             =   3435
      Width           =   3930
   End
   Begin MSChart20Lib.MSChart MSChart2 
      Height          =   2400
      Left            =   4470
      OleObjectBlob   =   "grafik.frx":4AF0
      TabIndex        =   1
      Top             =   420
      Width           =   3930
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   2400
      Left            =   135
      OleObjectBlob   =   "grafik.frx":6E48
      TabIndex        =   0
      Top             =   420
      Width           =   3930
   End
   Begin VB.Label Label4 
      Caption         =   "Hard Disk Perfonmance"
      Height          =   330
      Left            =   4875
      TabIndex        =   8
      Top             =   3015
      Width           =   3075
   End
   Begin VB.Label Label3 
      Caption         =   "Floating Point Perfonmace"
      Height          =   525
      Left            =   690
      TabIndex        =   7
      Top             =   3150
      Width           =   2745
   End
   Begin VB.Label Label2 
      Caption         =   "Graphics Perfonmance"
      Height          =   480
      Left            =   5130
      TabIndex        =   6
      Top             =   180
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Real Number Perfonmance"
      Height          =   420
      Left            =   780
      TabIndex        =   5
      Top             =   195
      Width           =   3045
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0C0C0&
      Height          =   6060
      Left            =   30
      Top             =   90
      Width           =   8700
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404040&
      Height          =   6120
      Left            =   0
      Top             =   75
      Width           =   8775
   End
End
Attribute VB_Name = "grafik"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
MSChart1.RowCount = 3
MSChart1.ColumnCount = 1
MSChart1.chartType = VtChChartType2dBar
MSChart1.Row = 1
MSChart1.Data = 228
MSChart1.RowLabel = "P II 233"
MSChart1.Row = 2
MSChart1.Data = 118
MSChart1.RowLabel = "686 - 120"
MSChart1.Row = 3
MSChart1.Data = 0
MSChart1.RowLabel = "Your CPU"

MSChart2.RowCount = 3
MSChart2.ColumnCount = 1
MSChart2.chartType = VtChChartType2dBar
MSChart2.Row = 1
MSChart2.Data = 27
MSChart2.RowLabel = "P II 233"
MSChart2.Row = 2
MSChart2.Data = 5.85
MSChart2.RowLabel = "686 - 120"
MSChart2.Row = 3
MSChart2.Data = 0
MSChart2.RowLabel = "Your CPU"

MSChart3.RowCount = 3
MSChart3.ColumnCount = 1
MSChart3.chartType = VtChChartType2dBar
MSChart3.Row = 1
MSChart3.Data = 7.33
MSChart3.RowLabel = "Trid 4MB AGP"
MSChart3.Row = 2
MSChart3.Data = 3
MSChart3.RowLabel = "1MB PCI"
MSChart3.Row = 3
MSChart3.Data = 0
MSChart3.RowLabel = "Your Graphics Card"

MSChart4.RowCount = 3
MSChart4.ColumnCount = 1
MSChart4.chartType = VtChChartType2dBar
MSChart4.Row = 1
MSChart4.Data = 2.6
MSChart4.RowLabel = "4.2 QUANTUM"
MSChart4.Row = 2
MSChart4.Data = 2.3
MSChart4.RowLabel = "2.1 SEAGATE"
MSChart4.Row = 3
MSChart4.Data = 0
MSChart4.RowLabel = "Your HD"


End Sub

Private Sub Command1_Click()
Dim i As Long, x As Long, t
t = Timer
Screen.MousePointer = 12
For i = 0 To 10000000
x = x + i + 10
x = 2 * i - 100 - x
Next
Screen.MousePointer = 0
t = Timer - t
MSChart1.Row = 3
MSChart1.Data = 100 / t

End Sub

Private Sub Command2_Click()
Dim i As Long, x As Long, t
t = Timer
Screen.MousePointer = 12
For i = 1 To 1000000
x = Log(i) * Sin(i) * Exp(Cos(i))
x = x / i
x = x * i
Next
Screen.MousePointer = 0
t = Timer - t
MSChart2.Row = 3
MSChart2.Data = 100 / t

End Sub
Private Sub Command3_Click()
Dim i As Long, x As Long, t
t = Timer
Screen.MousePointer = 12
For i = 1 To 20000
Line (Rnd * 1000, Rnd * 1000)-(Rnd * 1000, Rnd * 1000)
Line (Rnd * 1000, Rnd * 1000)-(Rnd * 1000, Rnd * 1000), Rnd * 100, BF
Circle (Rnd * 1000, Rnd * 1000), Rnd * 1000, Rnd * 1000
Next
Screen.MousePointer = 0
t = Timer - t
Cls
MSChart3.Row = 3
MSChart3.Data = 100 / t
End Sub
Private Sub command4_click()
Dim i As Long, x As String, t, driver
driver = Left(Drive1.Drive, 2)
t = Timer
Screen.MousePointer = 12
x = "tturrrkkkisshhhmaaaafffffiaa"
Open driver + "\tm.$$$" For Random As #1
For i = 1 To 150000
Put #1, i, x
Next
Close #1
Open driver + "\tm.$$$" For Random As #1
For i = 1 To 150000
Get #1, i, x
Next
Close #1
Open driver + "\tm.$$$" For Input As #1
While Not EOF(1)
Input #1, x
Wend
Close #1
x = "tturrrkkkisshhhmaaaafffffiaa"
Open driver + "\tm.$$$" For Output As #1
For i = 1 To 150000
Write #1, x
Next
Close #1
Kill driver + "\tm.$$$"
Screen.MousePointer = 0
t = Timer - t
MSChart4.Row = 3
MSChart4.Data = 100 / t

End Sub



























