VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{EC76FE26-BAFD-4E89-AA40-E748DA83A570}#1.0#0"; "IsButton_Ard.ocx"
Begin VB.Form Frm_Sel_Penj_Perkasir 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleksi Laporan Penj Perkasir"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6135
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
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   120
      ScaleHeight     =   2145
      ScaleWidth      =   5865
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      Begin MSComCtl2.DTPicker DTP1 
         Height          =   330
         Left            =   1440
         TabIndex        =   6
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         CalendarTitleBackColor=   12582912
         CalendarTitleForeColor=   16777215
         Format          =   22085633
         CurrentDate     =   39371
      End
      Begin MSComCtl2.DTPicker DTP2 
         Height          =   330
         Left            =   3360
         TabIndex        =   8
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         CalendarTitleBackColor=   12582912
         CalendarTitleForeColor=   16777215
         Format          =   22085633
         CurrentDate     =   39371
      End
      Begin IsButton_Ard.isButton Cmd_Tampil 
         Height          =   615
         Left            =   3360
         TabIndex        =   9
         Top             =   1440
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1085
         Icon            =   "Frm_Sel_Penj_Perkasir.frx":0000
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
      Begin IsButton_Ard.isButton Cmd_Keluar 
         Height          =   615
         Left            =   4560
         TabIndex        =   10
         Top             =   1440
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1085
         Icon            =   "Frm_Sel_Penj_Perkasir.frx":001C
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
         Caption         =   "s/d"
         Height          =   210
         Index           =   4
         Left            =   3000
         TabIndex        =   7
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Lbl_Nama 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1440
         TabIndex        =   5
         Top             =   240
         Width           =   4215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   210
         Index           =   3
         Left            =   1320
         TabIndex        =   4
         Top             =   600
         Width           =   60
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   210
         Index           =   2
         Left            =   1320
         TabIndex        =   3
         Top             =   240
         Width           =   60
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Periode"
         Height          =   210
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Kasir"
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   870
      End
   End
End
Attribute VB_Name = "Frm_Sel_Penj_Perkasir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cmd_Keluar_Click()
    Unload Me
End Sub

Private Sub Cmd_Tampil_Click()

On Error GoTo err_handler

Dim sql_tampil As String

sql_tampil = "select * from qr_penjualan_sebenarnya"
        
sql_tampil = sql_tampil & " where"

sql_tampil = sql_tampil & " tgl >= datevalue('" & Trim(DTP1.Value) & "') and tgl <= datevalue('" & Trim(DTP2.Value) & "')"

    If Lbl_Nama.Caption <> "" Then
        sql_tampil = sql_tampil & " and nama_user='" & Trim(Lbl_Nama.Caption) & "'"
    End If

sql_tampil = sql_tampil & " order by nama_user asc,no_faktur,tgl asc"

sqlku = sql_tampil

Load Frm_Lap_Penj_Perkasr
    Frm_Lap_Penj_Perkasr.Show
    
On Error GoTo 0
Exit Sub

err_handler:
    
    Dim konfirm As Integer
        konfirm = CInt(MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbInformation, "Information"))
            Err.Clear

End Sub

Private Sub DTP1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then DTP2.SetFocus
End Sub

Private Sub DTP2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Cmd_Tampil.SetFocus
End Sub

Private Sub Form_Activate()
On Error Resume Next
    DTP1.SetFocus
End Sub

Private Sub Form_Load()
    
    With Me
        .Left = Screen.Width / 2 - .Width / 2
        .Top = 350
    End With
    
    Lbl_Nama.Caption = utama.lbl_user.Caption
    DTP1.Value = Date
    DTP2.Value = Date
    
End Sub
