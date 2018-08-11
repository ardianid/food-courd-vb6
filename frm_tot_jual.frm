VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_tot_jual 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6255
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
   ScaleHeight     =   3150
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   120
      ScaleHeight     =   2865
      ScaleWidth      =   5985
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      Begin VB.CheckBox cek_tanggal 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Total  Pertanggal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3840
         TabIndex        =   9
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox txt_kode 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   3
         Top             =   1440
         Width           =   1455
      End
      Begin Crystal.CrystalReport rpt 
         Left            =   120
         Top             =   1680
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowTitle     =   "LAPORAN TOTAL PENJUALAN"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowGroupTree=   -1  'True
         WindowAllowDrillDown=   -1  'True
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin VB.CommandButton cmd_tampil 
         Caption         =   "Tampil"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4560
         TabIndex        =   4
         Top             =   2160
         Width           =   1215
      End
      Begin MSMask.MaskEdBox msk_tgl1 
         Height          =   375
         Left            =   2160
         TabIndex        =   1
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_tgl2 
         Height          =   375
         Left            =   4200
         TabIndex        =   2
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Counter"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   8
         Top             =   1440
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "s/d"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3720
         TabIndex        =   7
         Top             =   960
         Width           =   270
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   5640
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kriteria Pencetakan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   270
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   2220
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Periode"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   645
      End
   End
End
Attribute VB_Name = "frm_tot_jual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sql As String
Dim rs As New ADODB.Recordset

Private Sub Cmd_Tampil_Click()
    
On Error GoTo er_tampil
    
    Me.MousePointer = vbHourglass
    utama.MousePointer = vbHourglass
    
sql = ""

 If cek_tanggal.Value = vbUnchecked Then
    
    sql = sql & "select * from qr_penjualan_sebenarnya"
    
End If
    
If cek_tanggal.Value = vbChecked Then
    
    sql = "select * from qr_penjualan_sebenarnya"
    
End If
    
    If msk_tgl1.Text <> "__/__/____" And msk_tgl2.Text <> "__/__/____" Then
        sql = sql & " where tgl >= DateValue('" & Trim(msk_tgl1.Text) & "') and tgl <= DateValue('" & Trim(msk_tgl2.Text) & "')"
        
            If txt_kode.Text <> "" Then
                sql = sql & " and kode_counter='" & Trim(txt_kode.Text) & "'"
            End If
            
    End If
    
    sql = sql & " Order by kode_counter"
    
    sqlku = ""
    sqlku = sql
    
    If msk_tgl1.Text <> "__/__/____" Then
        tgl1 = msk_tgl1.Text
    Else
        tgl1 = ""
    End If
    
    If msk_tgl2.Text <> "__/__/____" Then
        tgl2 = msk_tgl2.Text
    Else
        tgl2 = ""
    End If
    
If cek_tanggal.Value = vbUnchecked Then
    Load Frm_Lap_Tot_Jual
        Frm_Lap_Tot_Jual.Show
End If

If cek_tanggal.Value = vbChecked Then
    Load Frm_Lap_Tot_Jual_Pertanggal
        Frm_Lap_Tot_Jual_Pertanggal.Show
End If

    
    If Me.MousePointer = vbHourglass Then
        Me.MousePointer = vbDefault
        utama.MousePointer = vbDefault
    End If
    
    Exit Sub
    
er_tampil:
    
    If Me.MousePointer = vbHourglass Then
        Me.MousePointer = vbDefault
        utama.MousePointer = vbDefault
    End If
    
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
    
End Sub

Private Sub Form_Activate()
    On Error Resume Next
        msk_tgl1.SetFocus
End Sub

Private Sub Form_Load()

With Me
    .Left = Screen.Width / 2 - .Width / 2
    .Top = 350
End With

End Sub

Private Sub msk_tgl1_GotFocus()
    msk_tgl1.SelStart = 0
    msk_tgl1.SelLength = Len(msk_tgl1)
End Sub

Private Sub msk_tgl1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then msk_tgl2.SetFocus
End Sub

Private Sub msk_tgl2_GotFocus()
    msk_tgl2.SelStart = 0
    msk_tgl2.SelLength = Len(msk_tgl2)
End Sub

Private Sub msk_tgl2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txt_kode.SetFocus
End Sub

Private Sub txt_kode_GotFocus()
    txt_kode.SelStart = 0
    txt_kode.SelLength = Len(txt_kode)
End Sub

Private Sub txt_kode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmd_tampil.SetFocus
End Sub
