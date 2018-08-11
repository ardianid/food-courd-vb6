VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_baker 
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8520
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   8295
      Left            =   240
      ScaleHeight     =   8265
      ScaleWidth      =   14625
      TabIndex        =   0
      Top             =   120
      Width           =   14655
      Begin VB.PictureBox pic_pegawai 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   6615
         Left            =   5880
         ScaleHeight     =   6585
         ScaleWidth      =   7065
         TabIndex        =   11
         Top             =   1320
         Visible         =   0   'False
         Width           =   7095
         Begin VB.CommandButton cmd_ok_peg 
            Caption         =   "Ok"
            Height          =   495
            Left            =   5520
            TabIndex        =   15
            Top             =   6000
            Width           =   1455
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00800000&
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   0
            ScaleHeight     =   345
            ScaleWidth      =   7065
            TabIndex        =   13
            Top             =   0
            Width           =   7095
            Begin VB.CommandButton cmd_x_pegawai 
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
               Left            =   6600
               TabIndex        =   14
               Top             =   0
               Width           =   495
            End
         End
         Begin TrueDBGrid60.TDBGrid grd_pegawai 
            Height          =   5175
            Left            =   120
            OleObjectBlob   =   "frm_baker.frx":0000
            TabIndex        =   12
            Top             =   720
            Width           =   6855
         End
      End
      Begin VB.PictureBox pic_jam 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   4815
         Left            =   2040
         ScaleHeight     =   4785
         ScaleWidth      =   5505
         TabIndex        =   7
         Top             =   480
         Visible         =   0   'False
         Width           =   5535
         Begin TrueDBGrid60.TDBGrid grd_jam 
            Height          =   4215
            Left            =   120
            OleObjectBlob   =   "frm_baker.frx":2CDD
            TabIndex        =   10
            Top             =   480
            Width           =   5295
         End
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
            TabIndex        =   9
            Top             =   0
            Width           =   495
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H00800000&
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   0
            ScaleHeight     =   345
            ScaleWidth      =   5505
            TabIndex        =   8
            Top             =   0
            Width           =   5535
         End
      End
      Begin VB.CommandButton cmd_hapus 
         Caption         =   "Hapus"
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
         Left            =   12960
         TabIndex        =   6
         Top             =   7680
         Width           =   1455
      End
      Begin VB.CommandButton cmd_Tambah 
         Caption         =   "Tambah"
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
         Left            =   11520
         TabIndex        =   5
         Top             =   7680
         Width           =   1455
      End
      Begin TrueDBGrid60.TDBGrid grd_daftar 
         Height          =   6735
         Left            =   240
         OleObjectBlob   =   "frm_baker.frx":60EE
         TabIndex        =   4
         Top             =   960
         Width           =   14175
      End
      Begin VB.CommandButton cmd_tampil 
         Caption         =   "Tampil"
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
         Left            =   12840
         TabIndex        =   3
         Top             =   480
         Width           =   1575
      End
      Begin MSMask.MaskEdBox msk_jam 
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
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
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1110
      End
   End
End
Attribute VB_Name = "frm_baker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arr_daftar As New XArrayDB
Dim arr_jam As New XArrayDB
Dim arr_pegawai As New XArrayDB
Dim id_jam As String, id_tugas As String

Private Sub kosong_pegawai()
    arr_pegawai.ReDim 0, 0, 0, 0
    grd_pegawai.ReBind
    grd_pegawai.Refresh
End Sub

Private Sub isi_pegawai()

On Error GoTo er_pegawai

    Dim sql, sql1 As String
    Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
    Dim a As Long
    Dim id_k, nama_k As String
        
        kosong_pegawai
        
        'sql = "select distinct(id) as id_k from qr_tugas where id_jam_kerja is null or id_jam_kerja <> " & id_jam
        'rs.Open sql, cn, adOpenKeyset
         sql = "select id,nama_karyawan from tbl_karyawan order by nama_karyawan"
         rs.Open sql, cn, adOpenKeyset
            If Not rs.EOF Then
                
                rs.MoveLast
                rs.MoveFirst
                
                a = 1
                    Do While Not rs.EOF
                      
                    'sql1 = "select nama_karyawan from tbl_karyawan where id=" & rs("id_k")
                    'rs1.Open sql1, cn
                    ' If Not rs1.EOF Then
                     
                        arr_pegawai.ReDim 1, a, 0, 4
                        grd_pegawai.ReBind
                        grd_pegawai.Refresh
                            
                            id_k = rs("id")
                            nama_k = rs("nama_karyawan")
                            
                        arr_pegawai(a, 0) = id_k
                        arr_pegawai(a, 1) = nama_k
                        arr_pegawai(a, 2) = vbUnchecked
                        
                     a = a + 1
                     
                     'End If
                    'rs1.Close
                    
                    rs.MoveNext
                    Loop
                    grd_pegawai.ReBind
                    grd_pegawai.Refresh
                    
                    
                    
            End If
        rs.Close
        grd_pegawai.MoveFirst
        
        Exit Sub
        
er_pegawai:
        Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
        
End Sub

Private Sub cmd_hapus_Click()
    
On Error GoTo er_hapus
    
    If arr_daftar.UpperBound(1) > 0 Then
        
        Dim sql, sql1 As String
        Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
            
            If MsgBox("Yakin akan hapus data karyawan " & arr_daftar(grd_daftar.Bookmark, 2), vbYesNo + vbQuestion, "Pesan") = vbNo Then
                Exit Sub
            End If
            
            sql = "select id from tr_pembagian_kerja where id=" & id_tugas
            rs.Open sql, cn
                If Not rs.EOF Then
                    sql1 = "delete from tr_pembagian_kerja where id=" & id_tugas
                    rs1.Open sql1, cn
                Else
                
                    MsgBox ("Data yang akan dihapus tidak ditemukan")
                End If
            rs.Close
            cmd_tampil_Click
    End If
    
    
    Exit Sub
    
er_hapus:
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
    
End Sub

Private Sub cmd_ok_peg_Click()
    Dim sql, sql1 As String
    Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
    Dim cek_dulu
    Dim a As Long
    
    On Error GoTo er_s
    
        cn.BeginTrans
        For a = 1 To arr_pegawai.UpperBound(1)
            cek_dulu = arr_pegawai(a, 2)
            
            If cek_dulu <> 0 Then
                sql1 = "select * from tr_pembagian_kerja where id_karyawan=" & arr_pegawai(a, 0) & " and id_jam_kerja=" & id_jam
                rs1.Open sql1, cn
                If rs1.EOF Then
                    sql = "insert into tr_pembagian_kerja (id_karyawan,id_jam_kerja) values(" & arr_pegawai(a, 0) & "," & id_jam & ")"
                    rs.Open sql, cn
                Else
                    MsgBox ("Nama karyawan " & arr_pegawai(a, 1) & " sudah ada pada jam kerja " & msk_jam.Text)
                    cn.RollbackTrans
                    Exit Sub
                End If
                rs1.Close
            End If
       Next a
       MsgBox ("Data berhasil disimpan")
       cn.CommitTrans
       pic_pegawai.Visible = False
       cmd_tampil_Click
       Exit Sub
       
er_s:
        cn.RollbackTrans
        Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
End Sub

Private Sub cmd_tambah_Click()
If msk_jam.Text <> "__:__:__" Then
    pic_pegawai.Visible = True
    pic_pegawai.Left = Me.Width / 2 - pic_pegawai.Width / 2
    pic_pegawai.Top = Me.Height / 2 - pic_pegawai.Height / 2
End If
End Sub

Private Sub cmd_tampil_Click()

On Error GoTo er_tampil

    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim idnya, namanya As String
    Dim a As Long
        
        If msk_jam.Text = "__:__:__" Then
            MsgBox ("Jam kerja harus diisi")
            msk_jam.SetFocus
            Exit Sub
        End If
        
        kosong_daftar
              
        sql = "select id,nama_karyawan from qr_pembagian_tugas where id_jam_kerja=" & id_jam
        rs.Open sql, cn, adOpenKeyset
            If Not rs.EOF Then
                
                rs.MoveLast
                rs.MoveFirst
                
                a = 1
                    Do While Not rs.EOF
                        arr_daftar.ReDim 1, a, 0, 3
                        grd_daftar.ReBind
                        grd_daftar.Refresh
                        
                            idnya = rs("id")
                            namanya = rs("nama_karyawan")
                                
                         arr_daftar(a, 0) = idnya
                         arr_daftar(a, 1) = a
                         arr_daftar(a, 2) = namanya
                    a = a + 1
                    rs.MoveNext
                    Loop
                    grd_daftar.ReBind
                    grd_daftar.Refresh
                    
            End If
         rs.Close
         
         isi_pegawai
         
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

Private Sub cmd_x_pegawai_Click()
    pic_pegawai.Visible = False
End Sub

Private Sub Form_Activate()
    msk_jam.SetFocus
End Sub

Private Sub Form_Load()
    
    grd_daftar.Array = arr_daftar
    
    grd_jam.Array = arr_jam
    
    grd_pegawai.Array = arr_pegawai
    
    isi_jam
    
    kosong_daftar
    
    Call cari_wewenang("Form Pembagian Tugas Kerja")
        
        If tambah_form = True Then
            cmd_tambah.Enabled = True
        Else
            cmd_tambah.Enabled = False
        End If
        
        
        If hapus_form = True Then
            cmd_hapus.Enabled = True
        Else
            cmd_hapus.Enabled = False
        End If
    
End Sub

Private Sub kosong_jam()
    arr_jam.ReDim 0, 0, 0, 0
    grd_jam.ReBind
    grd_jam.Refresh
End Sub

Private Sub isi_jam()

On Error GoTo er_jam

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
          
er_jam:
        Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
End Sub

Private Sub kosong_daftar()
    arr_daftar.ReDim 0, 0, 0, 0
    grd_daftar.ReBind
    grd_daftar.Refresh
End Sub

Private Sub grd_daftar_Click()
    On Error Resume Next
        If arr_daftar.UpperBound(1) > 0 Then
            id_tugas = arr_daftar(grd_daftar.Bookmark, 0)
        End If
End Sub

Private Sub grd_daftar_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    grd_daftar_Click
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

Private Sub grd_pegawai_AfterColUpdate(ByVal ColIndex As Integer)
 On Error GoTo er_saja
    
    If ColIndex = 2 Then
        arr_pegawai(grd_pegawai.Bookmark, ColIndex) = grd_pegawai.Columns(ColIndex).Text
    End If
    Exit Sub
    
er_saja:
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
End Sub

Private Sub pic_jam_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        pic_jam.Visible = False
        msk_jam.SetFocus
    End If
End Sub
