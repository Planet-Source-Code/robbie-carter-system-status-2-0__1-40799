VERSION 5.00
Object = "{1B773E42-2509-11CF-942F-008029004347}#3.6#0"; "sysmon.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loading..."
   ClientHeight    =   5370
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   7965
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   7965
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraSysMonitor 
      Caption         =   "&System Monitor"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      Begin SystemMonitorCtl.SystemMonitor SystemMonitor1 
         Height          =   4215
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   7575
         _Version        =   393219
         _ExtentX        =   13361
         _ExtentY        =   7435
         DisplayType     =   1
         ReportValueType =   0
         MaximumScale    =   100
         MinimumScale    =   0
         ShowLegend      =   -1  'True
         ShowToolbar     =   0   'False
         ShowScaleLabels =   -1  'True
         ShowHorizontalGrid=   0   'False
         ShowVerticalGrid=   0   'False
         ShowValueBar    =   0   'False
         ManualUpdate    =   0   'False
         Highlight       =   0   'False
         ReadOnly        =   -1  'True
         MonitorDuplicateInstances=   0   'False
         UpdateInterval  =   1
         DisplayFilter   =   1
         BackColorCtl    =   -2147483633
         ForeColor       =   -1
         BackColor       =   -1
         GridColor       =   8421504
         TimeBarColor    =   255
         Appearance      =   1
         BorderStyle     =   0
         NextCounterColor=   0
         NextCounterWidth=   0
         NextCounterLineStyle=   0
         GraphTitle      =   ""
         YAxisLabel      =   ""
         DataSourceType  =   1
         SqlDsnName      =   ""
         SqlLogSetName   =   ""
         LogFileCount    =   0
         AmbientFont     =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   "8.25"
         FontItalic      =   0   'False
         FontUnderline   =   0   'False
         FontStrikethrough=   0   'False
         FontWeight      =   400
         LegendColumnWidths=   $"frmMain.frx":0ECA
         LegendSortDirection=   0
         LegendSortColumn=   0
         CounterCount    =   0
         MaximumSamples  =   100
         SampleCount     =   0
      End
      Begin VB.CommandButton cmdSysMonitorDisplay 
         Caption         =   "Display Properties"
         Height          =   495
         Left            =   2040
         TabIndex        =   3
         Top             =   4560
         Width           =   1575
      End
      Begin VB.CommandButton cmdAddCounter 
         Caption         =   "Add Counter"
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   4560
         Width           =   1575
      End
   End
   Begin VB.Menu mnuProgram 
      Caption         =   "&Program"
      Begin VB.Menu mnuAlways 
         Caption         =   "Always On Top"
      End
      Begin VB.Menu mnuSeparator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuActions 
      Caption         =   "&Actions"
      Begin VB.Menu mnuMemory 
         Caption         =   "Memory"
         Begin VB.Menu mnuGetFull 
            Caption         =   "Get Full Memory Status"
         End
         Begin VB.Menu mnuGetQuick 
            Caption         =   "Get Quick Memory Status"
         End
      End
      Begin VB.Menu mnSysMon 
         Caption         =   "System Monitor"
         Begin VB.Menu mnuAddCounter 
            Caption         =   "Add Counter"
         End
         Begin VB.Menu mnuSysMonitorDisplay 
            Caption         =   "Display Properties"
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuMail 
         Caption         =   "Send e-mail to Developer"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright 2002 Bobby Carter

Private Sub cmdSysMonitorDisplay_Click()
SystemMonitor1.DisplayProperties
End Sub

Private Sub Form_Load()
Me.Caption = "System INFO v." & App.Major & "." & App.Minor & "." & App.Revision
End Sub
Private Sub mnuAbout_Click()
Me.Enabled = False
frmAbout.Show
End Sub

Private Sub mnuAddCounter_Click()
SystemMonitor1.BrowseCounters
End Sub

Private Sub mnuAlways_Click()
Call Always_On_Top(Me.hwnd, Me.Left / Screen.TwipsPerPixelX, _
Me.Top / Screen.TwipsPerPixelY, Me.Height / Screen.TwipsPerPixelY, Me.Width / Screen.TwipsPerPixelX, True)
mnuAlways.Checked = True
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub
Private Sub mnuGetFull_Click()
Me.Enabled = False
frmStatus.Show
End Sub
Private Sub mnuGetQuick_Click()
Dim s As Memorystatus
GlobalMemoryStatus s
TotalPhys = s.dwTotalPhys
MemoryLoad = s.dwMemoryLoad
AvailPhys = s.dwAvailPhys
info = info & "Total RAM  [MB]:" & TotalPhys \ 1048576 & vbCrLf & vbCrLf
info = info & "Unused RAM [MB]:" & AvailPhys \ 1048576 & vbCrLf
info = info & "Used RAM   [MB]:" & TotalPhys \ 1048576 - AvailPhys \ 1048576 & vbCrLf & vbCrLf
info = info & "Used RAM    [%]:" & MemoryLoad
MsgBox info, vbInformation, "Current RAM status"
End Sub
Private Sub mnuMail_Click()
Shell ("C:\Program Files\Outlook Express\msimn.exe /mailurl:mailto:sneak200@hotmail.com?")
End Sub

Private Sub mnuSysMonitorDisplay_Click()
SystemMonitor1.DisplayProperties
End Sub

Private Sub cmdAddCounter_Click()
SystemMonitor1.BrowseCounters
End Sub
