VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form driver 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Computer Information "
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6705
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   162
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "driver.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   6705
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   15
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "driver.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "driver.frx":2BF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "driver.frx":304A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6105
      Left            =   60
      TabIndex        =   0
      Top             =   75
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   10769
      _Version        =   393216
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BackColor       =   -2147483638
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Your Computer"
      TabPicture(0)   =   "driver.frx":349E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Shape1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Shape2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Shape3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "List1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "List2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "List3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Timer1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Command1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "List4"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Command2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Driver Information"
      TabPicture(1)   =   "driver.frx":34BA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "List5"
      Tab(1).Control(1)=   "Drive1"
      Tab(1).Control(2)=   "Label2"
      Tab(1).Control(3)=   "Label1"
      Tab(1).Control(4)=   "Image1"
      Tab(1).Control(5)=   "Shape4"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Windows Information"
      TabPicture(2)   =   "driver.frx":34D6
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "List6"
      Tab(2).Control(1)=   "Label6"
      Tab(2).Control(2)=   "Label5"
      Tab(2).Control(3)=   "Label4"
      Tab(2).Control(4)=   "Image2"
      Tab(2).Control(5)=   "Shape6"
      Tab(2).Control(6)=   "Shape5"
      Tab(2).Control(7)=   "Label3"
      Tab(2).Control(8)=   "Image3"
      Tab(2).ControlCount=   9
      Begin VB.ListBox List6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   660
         Left            =   -74835
         TabIndex        =   13
         Top             =   2565
         Width           =   3285
      End
      Begin VB.ListBox List5 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   4230
         Left            =   -74940
         TabIndex        =   10
         Top             =   1755
         Width           =   4020
      End
      Begin VB.DriveListBox Drive1 
         Height          =   330
         Left            =   -70755
         TabIndex        =   9
         Top             =   2325
         Width           =   2055
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Perfonmance Graphics"
         Height          =   870
         Left            =   3375
         TabIndex        =   6
         Top             =   5085
         Width           =   3165
      End
      Begin VB.ListBox List4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   870
         Left            =   60
         TabIndex        =   5
         Top             =   5085
         Width           =   3240
      End
      Begin VB.CommandButton Command1 
         Caption         =   "update"
         Height          =   690
         Left            =   135
         TabIndex        =   4
         Top             =   4335
         Width           =   2655
      End
      Begin VB.Timer Timer1 
         Left            =   855
         Top             =   3030
      End
      Begin VB.ListBox List3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   1710
         Left            =   120
         TabIndex        =   3
         Top             =   2610
         Width           =   2670
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   4440
         Left            =   2940
         TabIndex        =   2
         Top             =   495
         Width           =   3540
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   1920
         Left            =   135
         TabIndex        =   1
         Top             =   480
         Width           =   2640
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "KAYAHAN@PRESIDENCY.COM OR ICQ# 82533088"
         Height          =   585
         Left            =   -73860
         TabIndex        =   15
         Top             =   5295
         Width           =   4320
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "WINDOWS VERSION AND PATH"
         Height          =   270
         Left            =   -74805
         TabIndex        =   14
         Top             =   2280
         Width           =   3270
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "SETTINGS AND CONTROLS"
         Height          =   390
         Left            =   -73005
         TabIndex        =   12
         Top             =   1860
         Width           =   2715
      End
      Begin VB.Image Image2 
         Height          =   810
         Left            =   -74700
         Picture         =   "driver.frx":34F2
         Top             =   435
         Width           =   6015
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF0000&
         Height          =   870
         Left            =   -74700
         Shape           =   4  'Rounded Rectangle
         Top             =   480
         Width           =   6015
      End
      Begin VB.Shape Shape5 
         BorderWidth     =   3
         Height          =   675
         Left            =   -73785
         Top             =   1440
         Width           =   3840
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Microsoft Windows"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   -73725
         TabIndex        =   11
         Top             =   1485
         Width           =   3720
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Please select the drive that you want to see the detailed info of "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   -70905
         TabIndex        =   8
         Top             =   1740
         Width           =   2340
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "DRIVER INFO"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -74970
         TabIndex        =   7
         Top             =   1065
         Width           =   3960
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   -69720
         Picture         =   "driver.frx":3C13
         Top             =   735
         Width           =   480
      End
      Begin VB.Shape Shape4 
         Height          =   840
         Left            =   -70230
         Shape           =   4  'Rounded Rectangle
         Top             =   585
         Width           =   1530
      End
      Begin VB.Shape Shape3 
         Height          =   2505
         Left            =   75
         Top             =   2550
         Width           =   2760
      End
      Begin VB.Shape Shape2 
         Height          =   4620
         Left            =   2880
         Top             =   435
         Width           =   3660
      End
      Begin VB.Shape Shape1 
         Height          =   2070
         Left            =   90
         Top             =   435
         Width           =   2730
      End
      Begin VB.Image Image3 
         Height          =   5610
         Left            =   -74895
         Picture         =   "driver.frx":4055
         Stretch         =   -1  'True
         Top             =   390
         Width           =   6405
      End
   End
End
Attribute VB_Name = "driver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'need this custom type for later usage of GetDiskFreeSpaceEx declaration
'bu kisisel Type a daha sonra GetDiskFreeSpaceApiEx apisini kullanmak icin ihtiyacimiz olacak
Private Type uzunsayi
l As Long
h As Long
End Type
Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As uzunsayi, lpTotalNumberOfBytes As uzunsayi, lpTotalNumberOfFreeBytes As uzunsayi) As Long

'Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As uzun, lpBytesPerSector As uzun, lpNumberOfFreeClusters As uzun, lpTotalNumberOfClusters As Long) As Long
'the vb6 and vb5 api's(^) are different, but vb6 excepts the old one however I included the new one above also

Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long


Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal lpRootPathName As String) As Integer
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

Private Type MEMORYSTATUS
        dwLength As Long
        dwMemoryLoad As Long
        dwTotalPhys As Long
        dwAvailPhys As Long
        dwTotalPageFile As Long
        dwAvailPageFile As Long
        dwTotalVirtual As Long
        dwAvailVirtual As Long
End Type

Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)

Private Type SYSTEM_INFO
        dwOemID As Long
        dwPageSize As Long
        lpMinimumApplicationAddress As Long
        lpMaximumApplicationAddress As Long
        dwActiveProcessorMask As Long
        dwNumberOrfProcessors As Long
        dwProcessorType As Long
        dwAllocationGranularity As Long
        dwReserved As Long
End Type


Private Declare Function GetVersion Lib "kernel32" () As Long


Private Sub Command1_Click()
Timer1.Enabled = True
Timer1.Interval = 100
End Sub

Private Sub Command2_Click()
Load grafik
grafik.Show
End Sub

Private Sub Form_Load()
Show
Dim i As Integer, a As Integer
Dim s2, s3, s4, s5, s6
Dim x, uz
uz = 2 * List2.Width / Screen.TwipsPerPixelX
x = SendMessage(List2.hwnd, &H415, uz, 0)
uz = 2 * List4.Width / Screen.TwipsPerPixelX
x = SendMessage(List4.hwnd, &H415, uz, 0)
For i = 0 To 27
a = GetDriveType(Chr(65 + i) + ":\")
Select Case a
Case 0:
Case 2:
s2 = s2 + "  " + Chr(65 + i) + ":  "
Case 3:
s3 = s3 + "  " + Chr(65 + i) + ":  "
Case 4:
s4 = s4 + "  " + Chr(65 + i) + ":  "
Case 5:
s5 = s5 + "  " + Chr(65 + i) + ":  "
Case 6:
s6 = s6 + "  " + Chr(65 + i) + ":  "
End Select
Next
List1.AddItem ("          Drivers")
If s2 = "" Then List1.AddItem ("Disk Drivers    :   Is not Installed") Else List1.AddItem ("Disket Suruculer    :    " & s2)
If s3 = "" Then List1.AddItem ("Hard Disk Drivers:   Is not Installed") Else List1.AddItem ("Sabit Disk Suruculer:    " & s3)
If s4 = "" Then List1.AddItem ("Network Drivers   :   Is not Installed") Else List1.AddItem ("Network Suruculer   :    " & s4)
If s5 = "" Then List1.AddItem ("CDROM Drivers     :   Is not Installed") Else List1.AddItem ("CDROM Suruculer     :    " & s5)
If s6 = "" Then List1.AddItem ("RAM Disk            :   Is not Installed") Else List1.AddItem ("RAM Disk            :" & s6)
'by using systemmetrics , we can gather many of the information stored in our computer, Sytemmetrics i kullanarak
'bilgisayar hakkinda bir cok informasyonu bulabiliriz.
List2.AddItem ("                General System Information")
List2.AddItem ("Horizontal Screen Width:  " & GetSystemMetrics(0))
List2.AddItem ("Vertical Screen Width:  " & GetSystemMetrics(1))
List2.AddItem ("Arrow size on Horizantal Bars:  " & GetSystemMetrics(2))
List2.AddItem ("Arrow size on Vertical Bars  :  " & GetSystemMetrics(3))
List2.AddItem ("Height of the Active Title Bar:  " & GetSystemMetrics(4))
List2.AddItem ("Horizontal Width of Unsized window border: " & GetSystemMetrics(5))
List2.AddItem ("Vertical Width of Unsized window border:  " & GetSystemMetrics(6))
List2.AddItem ("Dialog Boxes Horizontal Border Thickness:  " & GetSystemMetrics(7))
List2.AddItem ("Dialog Boxes Vertical Border Thickness:  " & GetSystemMetrics(8))
List2.AddItem ("The Scroll Box height on Horizontal Scroll Bars:  " & GetSystemMetrics(9))
List2.AddItem ("The Scroll Box Width on Horizontal Scroll Bars:  " & GetSystemMetrics(10))
List2.AddItem ("Standard Icon Width:  " & GetSystemMetrics(11))
List2.AddItem ("Standard Icon Height:  " & GetSystemMetrics(12))
List2.AddItem ("Standard Cursor Width:  " & GetSystemMetrics(13))
List2.AddItem ("Standard Cursor Height:  " & GetSystemMetrics(14))
List2.AddItem ("Menu Height:  " & GetSystemMetrics(15))
List2.AddItem ("The inner Width of Maximized Window:  " & GetSystemMetrics(16))
List2.AddItem ("The inner Height of Maximized Window:  " & GetSystemMetrics(17))
If GetSystemMetrics(19) Then
List2.AddItem ("Mouse is connected to the computer  ")
Else
List2.AddItem ("Mouse is not connected to the computer  ")
End If
List2.AddItem ("The Arrow Heigth On Vertical Scroll Bar:  " & GetSystemMetrics(20))
List2.AddItem ("The Arrow Width On Vertical Scroll Bar:  " & GetSystemMetrics(21))
If GetSystemMetrics(22) Then
List2.AddItem ("Windows is working in Debug Mode   ")
Else
List2.AddItem ("Windows is not working in Debug Mode  ")
End If
If GetSystemMetrics(23) Then
List2.AddItem ("Mouse button functions are replaced   ")
Else
List2.AddItem ("Mouse button functions are not replaced     ")
End If
List2.AddItem ("The minumum width of opened windows:  " & GetSystemMetrics(28))
List2.AddItem ("The height of opened windows:  " & GetSystemMetrics(29))
List2.AddItem ("The Width of Picture on Active Window Border:  " & GetSystemMetrics(30))
List2.AddItem ("The Height of Picture on Active Window Border:  " & GetSystemMetrics(31))
List2.AddItem ("The Horizontal Deviance Width for Mouse Double Click:  " & GetSystemMetrics(36))
List2.AddItem ("The Horizontal Deviance Height for Mouse Double Click " & GetSystemMetrics(37))
List2.AddItem ("Desktop Icon Width:  " & GetSystemMetrics(38))
List2.AddItem ("Desktop Icon Height:  " & GetSystemMetrics(39))
If GetSystemMetrics(40) Then
List2.AddItem ("Popup Menus are opened to (directed to)the right")
List2.AddItem ("Popup Menus are not opened to (directed to)the right")
End If
If GetSystemMetrics(41) Then
List2.AddItem ("Pen is installed and its handle number:  " & GetSystemMetrics(41))
Else
List2.AddItem ("Pen is not installed.")
End If
If GetSystemMetrics(42) Then
List2.AddItem ("Windows is using 2 bytes for Charachter set (like chinese)")
Else
List2.AddItem ("Windows is not using 2 bytes for Charachter set (like chinese)")
End If
List2.AddItem ("Number of buttons on Mouse:  " & GetSystemMetrics(43))
List2.AddItem ("The width of Active Header (border):  " & GetSystemMetrics(51))
List2.AddItem ("The width of Maximized Windows:  " & GetSystemMetrics(61))
If GetSystemMetrics(63) Then
List2.AddItem ("Network is installed in your system")
Else
List2.AddItem ("Network is not installed in your system")
End If
Select Case GetSystemMetrics(67)
Case 0: List2.AddItem ("Opened Mode of Windows:  Normal")
Case 1: List2.AddItem ("Opened Mode of Windows:  Safe Mode")
Case 2: List2.AddItem ("Opened Mode of Windows:  Safe Mode with Network Support")
End Select
List2.AddItem ("The Width of Check Boxes in Menus:  " & GetSystemMetrics(71))
List2.AddItem ("The Height of Check Boxes in Menus:  " & GetSystemMetrics(72))
If GetSystemMetrics(73) Then
List2.AddItem ("Your computer is slow")
Else
List2.AddItem ("Your computer is at normal speed")
End If
Timer1.Interval = 100
Dim s As SYSTEM_INFO
GetSystemInfo s
List4.AddItem ("             CPU Info")

List4.AddItem (s.dwProcessorType & " type " & s.dwNumberOrfProcessors & " Proccesor")
List4.AddItem (s.lpMinimumApplicationAddress & " Kb  Minimum Application Address")
List4.AddItem (s.lpMaximumApplicationAddress & " Kb  Maximum Application Address")
List4.AddItem (s.dwActiveProcessorMask & " active processor has been masked")
drive1_change
Dim z
z = GetVersion()

List6.AddItem ("Windows  " & Str(z And &HFF) & "." & Str(a / &H100 And &HFF))
If z And &H7FFFFFFF Then
List6.AddItem ("System Win32s")
Else
List6.AddItem ("System Windows NT")

End If

End Sub





Private Sub Timer1_Timer()
Cls
Dim m As MEMORYSTATUS
GlobalMemoryStatus m
List3.Clear
List3.AddItem ("              MEMORY STATUS")
List3.AddItem ("Memory Usage:  %" & m.dwMemoryLoad)
List3.AddItem ("Total RAM:   " & m.dwTotalPhys / 1024 / 1024 & " MB")
List3.AddItem ("Empty (unused) RAM:    " & m.dwAvailPhys / 1024 / 1024 & " MB")
Timer1.Enabled = False
End Sub
Private Sub drive1_change()
Cls
Dim r As Long
Dim pos As Integer
Dim VolumeSN As Long
Dim MaxFNLen As Long
Dim FsName As String
Dim maxuz As Long
Dim flags As Long
Dim VolumeName As String
Dim surucu As String
VolumeName = Space(14)
FsName = Space(32)
surucu = Left(Drive1.Drive, 2) & "\"
Dim afb As uzunsayi, tb As uzunsayi, fb As uzunsayi
Call GetDiskFreeSpaceEx(surucu, afb, tb, fb)
List5.AddItem ("")
List5.AddItem ("             " & Drive1.Drive)
List5.AddItem ("Total size MB:  " & (tb.h * 2 ^ 32 + tb.l) / 1024 / 1024)
List5.AddItem ("Total Free Space MB:  " & (fb.h * 2 ^ 32 + fb.l) / 1024 / 1024)
List5.AddItem ("Total Free Space for Active User MB:  " & (afb.h * 2 ^ 32 + afb.l) / 1024 / 1024)
 r = GetVolumeInformation(surucu, VolumeName, Len(VolumeName), VolumeSN&, maxuz, flags, FsName, Len(FsName))
 If r = 0 Then
 Exit Sub
 End If
 pos = InStr(VolumeName, Chr(0))
 If pos Then
 VolumeName = Left(VolumeName, pos - 1)
 List5.AddItem ("Volume Name: " & VolumeName)
 List5.AddItem ("Serial No:   " & Hex(VolumeSN))
 List5.AddItem ("Max File Name Lenght :  " & maxuz)
End If
 pos = InStr(FsName, Chr(0))
 If pos Then
 FsName = Left(FsName, pos - 1)
 List5.AddItem ("System :  " & FsName)
 End If
 If flags And &H8000 Then
 List5.AddItem ("Compression :  " & " DblSpace or DrvSpace")
 End If
 If flags And &H10 Then
 List5.AddItem ("Compression :    " & " File Based")
 End If
 If flags And 1 Then
 List5.AddItem ("Differentiation of Upper Case & Lower Case Letters is Present")
 Else
 List5.AddItem ("Differentiation of Upper Case & Lower Case Letters is not Present")
 End If
 List5.AddItem ("-------------------------------------------------------")
End Sub

