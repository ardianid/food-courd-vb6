VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{EC76FE26-BAFD-4E89-AA40-E748DA83A570}#1.0#0"; "IsButton_Ard.ocx"
Begin VB.Form Frm_Sel_LapFaktur 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SELEKSI LAPORAN FAKTUR PER-PERIODE"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5550
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2310
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin IsButton_Ard.isButton Cmd_Tampil 
         Height          =   615
         Left            =   2760
         TabIndex        =   6
         Top             =   1320
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1085
         Icon            =   "Frm_Sel_LapFaktur.frx":0000
         Style           =   8
         Caption         =   "Tampil"
         iNonThemeStyle  =   0
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
      Begin MSComCtl2.DTPicker DTP_Tgl1 
         Height          =   375
         Left            =   1200
         TabIndex        =   3
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         CalendarTitleBackColor=   12582912
         CalendarTitleForeColor=   16777215
         Format          =   21954561
         CurrentDate     =   39213
      End
      Begin MSComCtl2.DTPicker DTP_Tgl2 
         Height          =   375
         Left            =   3360
         TabIndex        =   4
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         CalendarTitleBackColor=   12582912
         CalendarTitleForeColor=   16777215
         Format          =   21954561
         CurrentDate     =   39213
      End
      Begin IsButton_Ard.isButton Cmd_Keluar 
         Height          =   615
         Left            =   3960
         TabIndex        =   7
         Top             =   1320
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1085
         Icon            =   "Frm_Sel_LapFaktur.frx":001C
         Style           =   8
         Caption         =   "Keluar"
         iNonThemeStyle  =   0
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "S/D"
         Height          =   210
         Index           =   2
         Left            =   2880
         TabIndex        =   5
         Top             =   600
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   210
         Index           =   1
         Left            =   1080
         TabIndex        =   2
         Top             =   600
         Width           =   60
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   645
      End
   End
End
Attribute VB_Name = "Frm_Sel_LapFaktur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_keluar_Click()
    Unload Me
End Sub

Private Sub Cmd_Tampil_Click()
    
    sqlku = "select * from qr_penjualan_sebenarnya where tgl >= DateValue('" & Trim(DTP_Tgl1.Value) & "') and tgl <= DateValue('" & Trim(DTP_Tgl2.Value) & "')"
    
    Load Frm_Lap_PerFaktur
        Frm_Lap_PerFaktur.Show
    
End Sub

Private Sub Dtp_Tgl1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then DTP_Tgl2.SetFocus
End Sub

Private Sub Dtp_Tgl2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Cmd_Tampil.SetFocus
End Sub

Private Sub Form_Activate()
    On Error Resume Next
        DTP_Tgl1.SetFocus
End Sub

Private Sub Form_Load()
    
    With Me
        .Left = Screen.Width / 2 - .Width / 2
        .Top = 250
    End With
    
End Sub
