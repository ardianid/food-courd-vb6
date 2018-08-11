VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form Frm_Seting_Letak_Database 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Letak Database"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7230
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
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   5415
      Begin MSComDlg.CommonDialog cd 
         Left            =   840
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin VB.TextBox Txt_Letak 
         Height          =   1095
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   240
         Width           =   5055
      End
      Begin ucXPButton.XPButton cmd_browse 
         Height          =   495
         Left            =   5400
         TabIndex        =   1
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "&Browse"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Frm_Seting_Letak_Database.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmd_simpan 
         Height          =   495
         Left            =   5400
         TabIndex        =   2
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "&Simpan"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Frm_Seting_Letak_Database.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin ucXPButton.XPButton cmd_keluar 
      Height          =   495
      Left            =   5640
      TabIndex        =   4
      Top             =   1680
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Keluar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Frm_Seting_Letak_Database.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "Frm_Seting_Letak_Database"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cnc As Boolean
Dim Lt As String

Private Sub Cmd_Browse_Click()
On Error Resume Next
    
    Cnc = True
    
    With cd
        .CancelError = True
        .Filter = "Acces Database|*.dat;*.mdb"
        .ShowOpen
    End With
    
    Lt = cd.FileName
    
    Txt_Letak.Text = Lt
    cmd_simpan.SetFocus
    
End Sub

Private Sub Cmd_Keluar_Click()
    Cnc = False
    Unload Me
End Sub

Private Sub Cmd_Simpan_Click()
    
Dim konfirm As Integer
    
    If Txt_Letak.Text = "" Then
        konfirm = CInt(MsgBox("Lokasi database tidak ditemukan", vbOKOnly + vbInformation, "Informasi"))
        
        Cmd_Browse_Click
        Exit Sub
        
    End If
    
    If Set_Lokasi_Database(Lt) = True Then
        
        konfirm = CInt(MsgBox("Letak database berhasil disimpan", vbOKOnly + vbInformation, "Simpan"))
        Txt_Letak = Lokasi_Database
        
    End If
    
End Sub

Private Sub Form_Activate()
On Error Resume Next
    cmd_browse.SetFocus
End Sub

Private Sub Form_Load()
    
    Cnc = True
    
    Txt_Letak.Enabled = False
    
    
     If Lokasi_Database = 0 Then
        Txt_Letak = ""
     Else
        Txt_Letak = Lokasi_Database
     End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Cnc = True Then
        Cancel = True
    Else
        Cancel = False
    End If
End Sub

