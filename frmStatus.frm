VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmStatus 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Current Memory Status"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5115
   ControlBox      =   0   'False
   Icon            =   "frmStatus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   3
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy to clipboard"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   2520
      Width           =   1335
   End
   Begin RichTextLib.RichTextBox txtStatus 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   4048
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      MousePointer    =   99
      Appearance      =   0
      TextRTF         =   $"frmStatus.frx":0ECA
      MouseIcon       =   "frmStatus.frx":0F55
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
frmMain.Enabled = True
Unload Me
End Sub
Private Sub cmdCopy_Click()
Clipboard.SetText txtStatus.Text
End Sub
Private Sub cmdSave_Click()
CommonDialog1.Filter = "Rich Text Files (*.rtf)|*.rtf"
CommonDialog1.ShowSave
sfile = CommonDialog1.FileName
txtStatus.SaveFile sfile
End Sub
Private Sub Form_Load()
Dim s As Memorystatus
GlobalMemoryStatus s

TotalPhys = s.dwTotalPhys
MemoryLoad = s.dwMemoryLoad
AvailPhys = s.dwAvailPhys

PagMem = s.dwTotalPageFile
AvailPagMem = s.dwAvailPageFile

TotalVirtual = s.dwTotalVirtual
AvailVirtual = s.dwAvailVirtual

info = "SysStatus Report" & vbCrLf
info = info & "Date" & vbCrLf
info = info & Date & " " & Time & vbCrLf & vbCrLf
info = info & "Physical Memory           [MB]:" & TotalPhys \ 1048576 & vbCrLf
info = info & "Available Physical Memory [MB]:" & AvailPhys \ 1048576 & vbCrLf
info = info & "Used Physical Memory      [MB]:" & TotalPhys \ 1048576 - AvailPhys \ 1048576 & vbCrLf
info = info & "Used Physical Memory      [%]:" & MemoryLoad & vbCrLf & vbCrLf

info = info & "Paging File               [MB]:" & PagMem \ 1048576 & vbCrLf
info = info & "Available Paging File     [MB]:" & AvailPagMem \ 1048576 & vbCrLf
info = info & "Paging File Usage         [MB]:" & PagMem \ 1048576 - AvailPagMem \ 1048576 & vbCrLf & vbCrLf

info = info & "Virtual Memory            [MB]:" & TotalVirtual \ 1048576 & vbCrLf
info = info & "Available Virtual Memory  [MB]:" & AvailVirtual \ 1048576 & vbCrLf
info = info & "Used Virtual Memory       [MB]:" & TotalVirtual \ 1048576 - AvailVirtual \ 1048576 & vbCrLf

txtStatus.Text = info
End Sub
