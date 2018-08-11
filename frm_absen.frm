VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_absen 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pic_jam 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3975
      Left            =   2040
      ScaleHeight     =   3945
      ScaleWidth      =   5505
      TabIndex        =   9
      Top             =   960
      Visible         =   0   'False
      Width           =   5535
      Begin VB.CommandButton cmd_x 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   11
         Top             =   0
         Width           =   495
      End
      Begin TrueDBGrid60.TDBGrid grd_jam 
         Height          =   3375
         Left            =   120
         OleObjectBlob   =   "frm_absen.frx":0000
         TabIndex        =   10
         Top             =   480
         Width           =   5295
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         ScaleHeight     =   345
         ScaleWidth      =   5505
         TabIndex        =   12
         Top             =   0
         Width           =   5535
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6375
      Left            =   120
      ScaleHeight     =   6345
      ScaleWidth      =   9225
      TabIndex        =   6
      Top             =   1440
      Width           =   9255
      Begin VB.CommandButton cmd_simpan 
         Caption         =   "Simpan"
         Height          =   495
         Left            =   7800
         TabIndex        =   8
         Top             =   5760
         Width           =   1335
      End
      Begin TrueDBGrid60.TDBGrid grd_daftar 
         Height          =   5535
         Left            =   120
         OleObjectBlob   =   "frm_absen.frx":3411
         TabIndex        =   7
         Top             =   120
         Width           =   9015
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   120
      ScaleHeight     =   1185
      ScaleWidth      =   9225
      TabIndex        =   0
      Top             =   120
      Width           =   9255
      Begin VB.TextBox txt_karyawan 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   14
         Top             =   720
         Width           =   2415
      End
      Begin VB.CommandButton cmd_tampil 
         Caption         =   "Tampil"
         Height          =   495
         Left            =   7680
         TabIndex        =   5
         Top             =   600
         Width           =   1455
      End
      Begin MSMask.MaskEdBox msk_jam 
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSComCtl2.DTPicker dtp_tgl 
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   240
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   19660801
         CurrentDate     =   38629
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Karyawan"
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
         Left            =   3240
         TabIndex        =   13
         Top             =   720
         Width           =   1620
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jam Masuk"
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
         Left            =   480
         TabIndex        =   3
         Top             =   720
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl"
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
         Left            =   480
         TabIndex        =   1
         Top             =   240
         Width           =   300
      End
   End
End
Attribute VB_Name = "frm_absen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arr_jam As New XArrayDB
Dim arr_daftar As New XArrayDB
Dim id_jam As String

Private Sub cmd_simpan_Click()
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim a As Long, absen_sekarang
        
    On Error GoTo er_s
        
        If arr_daftar.UpperBound(1) > 0 Then
            cn.BeginTrans
            For a = 1 To arr_daftar.UpperBound(1)
                
                 
                sql = "insert into tr_absen (tgl,jam,id_pembagian_kerja,masuk_gak,ket,nama_user)"
                sql = sql & " values('" & Trim(dtp_tgl.Value) & "','" & utama.lbl_jam.Caption & "'," & arr_daftar(a, 0) & "," & rubah_apsen(arr_daftar(a, 3)) & ",'" & arr_daftar(a, 4) & "','" & utama.lbl_user.Caption & "')"
                rs.Open sql, cn
                
            Next a
            MsgBox ("Data berhasil disimpan")
            cn.CommitTrans
            msk_jam.Text = "__:__:__"
            kosong_daftar
            msk_jam.SetFocus
            Exit Sub
        End If
        Exit Sub
            
er_s:
       Dim psn
        cn.RollbackTrans
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
        
End Sub

Private Sub cmd_tampil_Click()

On Error GoTo er_tampil

    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim i, n As String
    Dim a As Long
        
        kosong_daftar
        
        If msk_jam.Text = "__:__:__" Then
            MsgBox ("Jam Masuk Hrs Diisi")
            msk_jam.SetFocus
            Exit Sub
        End If
        
        If txt_karyawan.Text = "" Then
            txt_karyawan.Text = "Semua"
        End If
        
        If msk_jam.Text <> "__:__:__" And txt_karyawan.Text = "Semua" Then
            sql = "select id,nama_karyawan from qr_pembagian_tugas where id_jam_kerja=" & id_jam
        ElseIf msk_jam.Text <> "__:__:__" And txt_karyawan.Text <> "Semua" Then
            sql = "select id,nama_karyawan from qr_pembagian_tugas where nama_karyawan='" & Trim(txt_karyawan.Text) & "' and id_jam_kerja=" & id_jam
        End If
        
            rs.Open sql, cn, adOpenKeyset
                If Not rs.EOF Then
                    
                    rs.MoveLast
                    rs.MoveFirst
                    
                  a = 1
                    Do While Not rs.EOF
                        arr_daftar.ReDim 1, a, 0, 5
                        grd_daftar.ReBind
                        grd_daftar.Refresh
                            
                            i = rs("id")
                            n = rs("nama_karyawan")
                            
                        arr_daftar(a, 0) = i
                        arr_daftar(a, 1) = a
                        arr_daftar(a, 2) = n
                        arr_daftar(a, 3) = "Hadir"
                        arr_daftar(a, 4) = "-"
                    a = a + 1
                    rs.MoveNext
                    Loop
                    grd_daftar.ReBind
                    grd_daftar.Refresh
                End If
            rs.Close
            grd_daftar.MoveFirst
            
    Exit Sub
    
er_tampil:
        Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
            
End Sub

Private Sub cmd_x_Click()
    pic_jam.Visible = False
    msk_jam.SetFocus
End Sub

Private Sub Form_Load()

grd_daftar.Array = arr_daftar

grd_jam.Array = arr_jam

kosong_daftar

dtp_tgl.Value = Format(Date, "dd/mm/yyyy")

isi_jam

txt_karyawan.Text = "Semua"

Me.Left = utama.Width / 2 - Me.Width / 2
Me.Top = utama.Height / 2 - Me.Height / 2 - 1350

End Sub

Private Sub kosong_daftar()
    arr_daftar.ReDim 0, 0, 0, 0
    grd_daftar.ReBind
    grd_daftar.Refresh
End Sub

Private Sub grd_daftar_AfterColUpdate(ByVal ColIndex As Integer)

On Error GoTo er_after

    If ColIndex = 3 Then
        arr_daftar(grd_daftar.Bookmark, ColIndex) = grd_daftar.Columns(ColIndex).Text
    End If
    If ColIndex = 4 Then
        arr_daftar(grd_daftar.Bookmark, ColIndex) = grd_daftar.Columns(ColIndex).Text
    End If
    
    Exit Sub
    
er_after:
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
    
End Sub

Private Sub msk_jam_GotFocus()
    msk_jam.SelStart = 0
    msk_jam.SelLength = Len(msk_jam)
End Sub

Private Sub msk_jam_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        msk_jam.Text = "__:__:__"
        pic_jam.Visible = True
        grd_jam.SetFocus
    End If
    
    If KeyCode = 13 Then
        cmd_tampil.SetFocus
    End If
End Sub

Private Sub msk_jam_LostFocus()

On Error GoTo er_lost

    If msk_jam.Text <> "__:__:__" Then
        Dim sql As String
        Dim rs As New ADODB.Recordset
            
            sql = "select id from tbl_jam_kerja where jam_masuk=timevalue('" & Trim(msk_jam.Text) & "')"
            rs.Open sql, cn
                If Not rs.EOF Then
                    id_jam = rs("id")
                Else
                    MsgBox ("Jam kerja yang anda masukkan tidak ditemukan")
                    msk_jam.SetFocus
                End If
            rs.Close
    End If
    
    Exit Sub
    
er_lost:
        Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
End Sub

Private Sub kosong_jam()
    arr_jam.ReDim 0, 0, 0, 0
    grd_jam.ReBind
    grd_jam.Refresh
End Sub

Private Sub isi_jam()

On Error GoTo er_isi_jam

    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim jam_masuk, jam_is, jam_pulang, id_dia As String
    Dim a As Long
            
        kosong_jam
            
        sql = "select id,jam_masuk,jam_istirahat,jam_pulang from tbl_jam_kerja order by jam_masuk"
        rs.Open sql, cn, adOpenKeyset
            If Not rs.EOF Then
                
                rs.MoveLast
                rs.MoveFirst
                
            a = 1
                Do While Not rs.EOF
                    arr_jam.ReDim 1, a, 0, 5
                    grd_jam.ReBind
                    grd_jam.Refresh
                    
                    id_dia = rs("id")
                    
                    If Not IsNull(rs("jam_masuk")) Then
                        jam_masuk = Format(rs("jam_masuk"), "hh:mm:ss")
                    Else
                        jam_masuk = ""
                    End If
                    
                    If Not IsNull(rs("jam_istirahat")) Then
                        jam_is = Format(rs("jam_istirahat"), "hh:mm:ss")
                    Else
                        jam_is = ""
                    End If
                    
                    If Not IsNull(rs("jam_pulang")) Then
                        jam_pulang = Format(rs("jam_pulang"), "hh:mm:ss")
                    Else
                        jam_pulang = ""
                    End If
                    
                arr_jam(a, 0) = id_dia
                arr_jam(a, 1) = jam_masuk
                arr_jam(a, 2) = jam_is
                arr_jam(a, 3) = jam_pulang
                
                a = a + 1
                rs.MoveNext
                Loop
                grd_jam.ReBind
                grd_jam.Refresh
            End If
          rs.Close
                    
          Exit Sub
          
er_isi_jam:
        Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
                    
End Sub

Private Sub grd_jam_Click()
    On Error Resume Next
        If arr_jam.UpperBound(1) > 0 Then
            id_jam = arr_jam(grd_jam.Bookmark, 0)
        End If
End Sub

Private Sub grd_jam_DblClick()
    If arr_jam.UpperBound(1) > 0 Then
        msk_jam.Text = arr_jam(grd_jam.Bookmark, 1)
        pic_jam.Visible = False
        msk_jam.SetFocus
    End If
End Sub

Private Sub grd_jam_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And arr_jam.UpperBound(1) > 0 Then
        grd_jam_DblClick
    End If
    
    If KeyCode = vbKeyEscape Then
        pic_jam.Visible = False
        msk_jam.SetFocus
    End If
    
End Sub

Private Sub grd_jam_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    grd_jam_Click
End Sub

Private Sub pic_jam_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        pic_jam.Visible = False
        msk_jam.SetFocus
    End If
End Sub
