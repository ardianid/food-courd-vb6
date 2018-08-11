VERSION 5.00
Begin VB.Form frm_backup 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   9735
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pic1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5295
      Left            =   4200
      ScaleHeight     =   5265
      ScaleWidth      =   5025
      TabIndex        =   13
      Top             =   1680
      Visible         =   0   'False
      Width           =   5055
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   3360
         ScaleHeight     =   585
         ScaleWidth      =   1425
         TabIndex        =   20
         Top             =   120
         Width           =   1455
         Begin VB.CommandButton cmd_drive 
            Caption         =   "OK"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   0
            TabIndex        =   21
            Top             =   0
            Width           =   1455
         End
      End
      Begin VB.PictureBox Picture6 
         Height          =   495
         Left            =   120
         ScaleHeight     =   435
         ScaleWidth      =   4635
         TabIndex        =   18
         Top             =   4680
         Width           =   4695
         Begin VB.Label lbl_drive 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label6"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   120
            Width           =   645
         End
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   3735
         Left            =   120
         ScaleHeight     =   3705
         ScaleWidth      =   4665
         TabIndex        =   16
         Top             =   840
         Width           =   4695
         Begin VB.DirListBox Dir1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3405
            Left            =   120
            TabIndex        =   17
            Top             =   120
            Width           =   4335
         End
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   120
         ScaleHeight     =   585
         ScaleWidth      =   3105
         TabIndex        =   14
         Top             =   120
         Width           =   3135
         Begin VB.DriveListBox Drive1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   15
            Top             =   120
            Width           =   2775
         End
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      ScaleHeight     =   705
      ScaleWidth      =   9465
      TabIndex        =   10
      Top             =   3240
      Width           =   9495
      Begin VB.CommandButton cmd_proses 
         Caption         =   "Proses"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7920
         TabIndex        =   11
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   120
      ScaleHeight     =   2985
      ScaleWidth      =   9465
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      Begin VB.CommandButton cmd_browse 
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8760
         TabIndex        =   12
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txt_nama 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3480
         TabIndex        =   9
         Top             =   2400
         Width           =   5175
      End
      Begin VB.TextBox txt_letak 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3480
         TabIndex        =   7
         Top             =   1560
         Width           =   5175
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Left            =   240
         ScaleHeight     =   2115
         ScaleWidth      =   2595
         TabIndex        =   2
         Top             =   600
         Width           =   2655
         Begin VB.Label Label4 
            BackColor       =   &H8000000E&
            Caption         =   "Sebelum melakukan Backup Database pastikan tidak ada user lain yang sedang memakai program ini"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1455
            Left            =   120
            TabIndex        =   4
            Top             =   600
            Width           =   2295
         End
         Begin VB.Line Line2 
            BorderWidth     =   2
            X1              =   240
            X2              =   2400
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Informasi"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   3
            Top             =   120
            Width           =   1080
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama database backup"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3120
         TabIndex        =   8
         Top             =   2040
         Width           =   2505
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Letak folder hasil backup"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3120
         TabIndex        =   6
         Top             =   1200
         Width           =   2595
      End
      Begin VB.Label lbl_letak 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lbl_letak"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3120
         TabIndex        =   5
         Top             =   600
         Width           =   6015
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   9120
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Backup Database"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   345
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   2430
      End
   End
End
Attribute VB_Name = "frm_backup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim letaknya As String

Private Sub cmd_browse_Click()
    
    penuh
    
    pic1.Visible = True
    
    Drive1.SetFocus
    
End Sub

Private Sub cmd_drive_Click()
 
 normal
 
 If lbl_drive.Caption <> "" Then
        
        Dim cari_miring, drive_sekarang As String
            
            cari_miring = Right(Trim(lbl_drive.Caption), 1)
            
            If cari_miring <> "\" Then
                drive_sekarang = Trim(lbl_drive.Caption) & "\"
            Else
                drive_sekarang = Trim(lbl_drive.Caption)
            End If
            
            txt_letak.Text = drive_sekarang
            
 End If
    
 pic1.Visible = False
 
 txt_letak.SetFocus
 
End Sub

Private Sub cmd_proses_Click()
    
On Error GoTo er_proses
    
    If txt_letak.Text = "" Then MsgBox ("Letak folder harus diisi"): txt_letak.SetFocus: Exit Sub
    If txt_nama.Text = "" Then MsgBox ("Nama database hasil backup harus diisi"): txt_nama.SetFocus: Exit Sub
    
    Dim pindahkan
    Dim letak_baru, nama_baru As String
        
        nama_baru = Trim(txt_nama.Text) & ".mdb"
        
        letak_baru = Trim(txt_letak.Text) & nama_baru
        
        pindahkan = CopyFile(lbl_letak.Caption, letak_baru, 0)
    
    Dim benar As Integer
        benar = MsgBox("Proses backup berhasil dilakukan", vbOKOnly + vbInformation, App.Title)
        
        On Error GoTo 0
        
        Exit Sub
        
er_proses:
        Dim psn As Integer
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, App.Title)
            Err.Clear
    
End Sub

Private Sub Dir1_Change()

On Error Resume Next
    
    letaknya = ""
    
    letaknya = Dir1.path
    
    lbl_drive.Caption = ""
    
    lbl_drive.Caption = letaknya
    
End Sub

Private Sub Drive1_Change()

On Error GoTo er_drive

    Dir1.path = Drive1.Drive
    
    letaknya = ""
    
    letaknya = Drive1.Drive
    
    lbl_drive.Caption = ""
    
    lbl_drive.Caption = letaknya
    
    On Error GoTo 0
    
    Exit Sub
    
er_drive:
        
        Dim psn As Integer
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, App.Title)
            Err.Clear
    
End Sub

Private Sub Form_Load()
    lbl_letak.Caption = path
    
    Dim tanggal, bulan  As Long, tahun As String
        
        tanggal = DatePart("d", Now)
        
        If Len(tanggal) = 1 Then tanggal = 0 & tanggal
            
        bulan = DatePart("m", Now)
        
        tahun = DatePart("yyyy", Now)
        
        tahun = Right(tahun, 2)
        
        txt_nama.Text = tanggal & bulan & tahun
        
        letaknya = ""
        
        lbl_drive.Caption = letaknya
        
        normal
        
    Me.Left = Screen.Width / 2 - Me.Width / 2
    Me.Top = Screen.Height / 2 - Me.Height / 2 - 2900
    
End Sub

Sub normal()
    Me.Height = 4575
    Me.Width = 9825
    Me.ScaleHeight = 7095
    Me.ScaleWidth = 9735
End Sub

Sub penuh()
    Me.Height = 7575
    Me.Width = 9825
    Me.ScaleHeight = 7095
    Me.ScaleWidth = 9735
End Sub

Private Sub txt_letak_GotFocus()
    txt_letak.SelStart = 0
    txt_letak.SelLength = Len(txt_letak)
End Sub

Private Sub txt_nama_GotFocus()
    txt_nama.SelStart = 0
    txt_nama.SelLength = Len(txt_nama)
End Sub
