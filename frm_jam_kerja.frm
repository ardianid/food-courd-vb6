VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_jam_kerja 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   9675
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pic_tambah 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   120
      ScaleHeight     =   2865
      ScaleWidth      =   6945
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   6975
      Begin VB.CommandButton cmd_cancel 
         Caption         =   "Cancel"
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
         Left            =   5640
         TabIndex        =   14
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CommandButton cmd_simpan 
         Caption         =   "Simpan"
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
         Left            =   4320
         TabIndex        =   9
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox txt_ket 
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
         Left            =   1800
         TabIndex        =   8
         Top             =   1680
         Width           =   4815
      End
      Begin MSMask.MaskEdBox msk_masuk 
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
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
      Begin MSMask.MaskEdBox msk_istirahat 
         Height          =   375
         Left            =   1800
         TabIndex        =   6
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
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
      Begin MSMask.MaskEdBox msk_pulang 
         Height          =   375
         Left            =   1800
         TabIndex        =   7
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ket"
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
         Left            =   240
         TabIndex        =   4
         Top             =   1680
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jam Pulang"
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
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   1230
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jam Istirahat"
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
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jam Masuk"
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
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1200
      End
   End
   Begin VB.PictureBox pic_daftar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6255
      Left            =   120
      ScaleHeight     =   6225
      ScaleWidth      =   9465
      TabIndex        =   10
      Top             =   120
      Width           =   9495
      Begin TrueDBGrid60.TDBGrid grd_daftar 
         Height          =   5415
         Left            =   120
         OleObjectBlob   =   "frm_jam_kerja.frx":0000
         TabIndex        =   15
         Top             =   120
         Width           =   9255
      End
      Begin VB.CommandButton cmd_tambah 
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
         Left            =   5160
         TabIndex        =   13
         Top             =   5640
         Width           =   1335
      End
      Begin VB.CommandButton cmd_edit 
         Caption         =   "Edit"
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
         Left            =   6600
         TabIndex        =   12
         Top             =   5640
         Width           =   1335
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
         Left            =   8040
         TabIndex        =   11
         Top             =   5640
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frm_jam_kerja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arr_daftar As New XArrayDB
Dim id_jam As String
Dim sqla As String
Dim tujuan As Boolean

Private Sub kosong_daftar()
    arr_daftar.ReDim 0, 0, 0, 0
    grd_daftar.ReBind
    grd_daftar.Refresh
End Sub

Private Sub isi()
    
On Error GoTo er_isi
    
    Dim rs As New ADODB.Recordset
        
        kosong_daftar
            
        sqla = "select * from tbl_jam_kerja"
        rs.Open sqla, cn, adOpenKeyset
            If Not rs.EOF Then
                
                rs.MoveLast
                rs.MoveFirst
                
                lanjut rs
            End If
        rs.Close
            
        Exit Sub
        
er_isi:
        Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
            
End Sub

Private Sub lanjut(rs As Recordset)
    Dim a As Long
    Dim i_a, jm_msk, jm_is, jm_plg, ket As String
        
        a = 1
            Do While Not rs.EOF
                arr_daftar.ReDim 1, a, 0, 7
                grd_daftar.ReBind
                grd_daftar.Refresh
                    
                    i_a = rs("id")
                    If Not IsNull(rs("jam_masuk")) Then
                        jm_msk = Format(rs("jam_masuk"), "hh:mm:ss")
                    Else
                        jm_msk = ""
                    End If
                    
                    If Not IsNull(rs("jam_istirahat")) Then
                        jm_is = Format(rs("jam_istirahat"), "hh:mm:ss")
                    Else
                        jm_is = ""
                    End If
                    
                    If Not IsNull(rs("jam_pulang")) Then
                        jm_plg = Format(rs("jam_pulang"), "hh:mm:ss")
                    Else
                        jm_plg = ""
                    End If
                    
                    If Not IsNull(rs("ket")) Then
                        ket = rs("ket")
                    Else
                        ket = ""
                    End If
                    
               arr_daftar(a, 0) = i_a
               arr_daftar(a, 1) = a
               arr_daftar(a, 2) = jm_msk
               arr_daftar(a, 3) = jm_is
               arr_daftar(a, 4) = jm_plg
               arr_daftar(a, 5) = ket
               
          a = a + 1
          rs.MoveNext
          Loop
          grd_daftar.ReBind
          grd_daftar.Refresh
               
                    
End Sub

Private Sub cmd_cancel_Click()
    penuh
    isi
End Sub

Private Sub cmd_edit_Click()

On Error GoTo er_edit

    tujuan = False
    setengah
    
    If arr_daftar(grd_daftar.Bookmark, 2) <> "" Then
        msk_masuk.Text = Format(arr_daftar(grd_daftar.Bookmark, 2), "hh:mm:ss")
    Else
        msk_masuk.Text = "__:__:__"
    End If
    
    If arr_daftar(grd_daftar.Bookmark, 3) <> "" Then
        msk_istirahat.Text = Format(arr_daftar(grd_daftar.Bookmark, 3), "hh:mm:ss")
    Else
        msk_istirahat.Text = "__:__:__"
    End If
    
    If arr_daftar(grd_daftar.Bookmark, 4) <> "" Then
        msk_pulang.Text = Format(arr_daftar(grd_daftar.Bookmark, 4), "hh:mm:ss")
    Else
        msk_pulang.Text = "__:__:__"
    End If
    
    txt_ket.Text = arr_daftar(grd_daftar.Bookmark, 5)
    
    msk_masuk.SetFocus
    
    Exit Sub
    
er_edit:
    Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
    
End Sub

Private Sub cmd_hapus_Click()
    Dim sql, sql1 As String
    Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
    
    On Error GoTo er_hapus
    
    If arr_daftar.UpperBound(1) = 0 Then
        Exit Sub
    End If
    
    If MsgBox("Yakin akan hapus jam masuk " & arr_daftar(grd_daftar.Bookmark, 2), vbYesNo + vbQuestion, "Pesan") = vbNo Then
        Exit Sub
    End If
    
    sql = "select id from tbl_jam_kerja where id=" & id_jam
    rs.Open sql, cn
        If Not rs.EOF Then
            
            sql1 = "delete from tbl_jam_kerja where id=" & id_jam
            rs1.Open sql1, cn
            
        Else
            
            MsgBox ("Data yang akan dihapus tidak ditemukan")
            
        End If
    rs.Close
    isi
    Exit Sub
    
er_hapus:
        Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
    
End Sub

Private Sub cmd_simpan_Click()
    Dim sql, sql1 As String
    Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
        
On Error GoTo er_simpan
        
        If msk_masuk.Text = "__:__:__" Then
            MsgBox ("Jam masuk harus diisi")
        End If
        
        Dim jam_istirahat, jam_pulang
                    If msk_istirahat.Text = "__:__:__" Then
                        jam_istirahat = ""
                    Else
                        jam_istirahat = Trim(msk_istirahat.Text)
                    End If
                    
                    If msk_pulang.Text = "__:__:__" Then
                        jam_pulang = ""
                    Else
                        jam_pulang = Trim(msk_pulang.Text)
                    End If
        
     If tujuan = True Then
        sql = "select jam_masuk from tbl_jam_kerja where jam_masuk = timevalue('" & Trim(msk_masuk.Text) & "')"
        rs.Open sql, cn
            If rs.EOF Then
                    
                    sql1 = "insert into tbl_jam_kerja (jam_masuk,jam_istirahat,jam_pulang,ket) values('" & Trim(msk_masuk.Text) & "','" & jam_istirahat & "','" & jam_pulang & "','" & Trim(txt_ket.Text) & "')"
                    rs1.Open sql1, cn
                    
                    MsgBox ("Data berhasil disimpan")
                    kosong
                    msk_masuk.SetFocus
            Else
                    MsgBox ("Jam masuk sudah ada")
            End If
        rs.Close
     ElseIf tujuan = False Then
            sql = "select id from tbl_jam_kerja where id=" & id_jam
            rs.Open sql, cn
                If Not rs.EOF Then
                    sql1 = "update tbl_jam_kerja set jam_masuk='" & Trim(msk_masuk.Text) & "',jam_istirahat='" & jam_istirahat & "',jam_pulang='" & jam_pulang & "',ket='" & Trim(txt_ket.Text) & "' where id=" & id_jam
                    rs1.Open sql1, cn
                    
                    MsgBox ("Data berhasil diedit")
                Else
                    MsgBox ("Data yang akan diedit tidak ditemukan")
                End If
           rs.Close
           cmd_cancel_Click
     End If
    Exit Sub
        
er_simpan:
        Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
End Sub

Private Sub kosong()
    msk_masuk.Text = "__:__:__"
    msk_istirahat.Text = "__:__:__"
    msk_pulang.Text = "__:__:__"
    txt_ket.Text = ""
End Sub

Private Sub cmd_tambah_Click()
    tujuan = True
    setengah
    kosong
    msk_masuk.SetFocus
End Sub

Private Sub Form_Load()

    grd_daftar.Array = arr_daftar
    
    isi
    
    penuh
    
    
    Call cari_wewenang("Form Data Jam Kerja")
        
        If tambah_form = True Then
            cmd_tambah.Enabled = True
        Else
            cmd_tambah.Enabled = False
        End If
        
        If edit_form = True Then
            cmd_edit.Enabled = True
        Else
            cmd_edit.Enabled = False
        End If
        
        If hapus_form = True Then
            cmd_hapus.Enabled = True
        Else
            cmd_hapus.Enabled = False
        End If
        
    Me.Left = utama.Width / 2 - Me.Width / 2
    Me.Top = utama.Height / 2 - Me.Height / 2 - 2000
    
End Sub

Private Sub grd_daftar_Click()
    On Error Resume Next
        If arr_daftar.UpperBound(1) > 0 Then
            id_jam = arr_daftar(grd_daftar.Bookmark, 0)
        End If
End Sub

Private Sub grd_daftar_HeadClick(ByVal ColIndex As Integer)

On Error GoTo er_head

    Dim rs As New ADODB.Recordset
    Dim sql As String
    
    If sqla = "" Then
        Exit Sub
    End If
    
If arr_daftar.UpperBound(1) > 0 Then
    
    sql = ""
    sql = sql & sqla
    
    Select Case ColIndex
        
        Case 2
            sql = sql & " order by jam_masuk"
        Case 3
            sql = sql & " order by jam_istirahat"
        Case 4
            sql = sql & " order by jam_pulang"
        Case 5
            sql = sql & " order by ket"
   End Select
    
   rs.Open sql, cn, adOpenKeyset
        If Not rs.EOF Then
            
            rs.MoveLast
            rs.MoveFirst
            
            lanjut rs
        End If
   rs.Close
   
End If

Exit Sub

er_head:
    Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear

End Sub

Private Sub grd_daftar_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    grd_daftar_Click
End Sub

Private Sub msk_istirahat_GotFocus()
    msk_istirahat.SelStart = 0
    msk_istirahat.SelLength = Len(msk_istirahat)
End Sub

Private Sub msk_masuk_GotFocus()
    msk_masuk.SelStart = 0
    msk_masuk.SelLength = Len(msk_masuk)
End Sub
Private Sub msk_pulang_GotFocus()
    msk_pulang.SelStart = 0
    msk_pulang.SelLength = Len(msk_pulang)
End Sub

Private Sub txt_ket_GotFocus()
    txt_ket.SelStart = 0
    txt_ket.SelLength = Len(txt_ket)
End Sub

Private Sub penuh()
    pic_daftar.Visible = True
    pic_tambah.Visible = False
    Me.Height = 6975
    Me.Width = 9825
    Me.ScaleHeight = 6495
    Me.ScaleWidth = 9735
End Sub

Private Sub setengah()
    pic_daftar.Visible = False
    pic_tambah.Visible = True
    Me.Height = 3615
    Me.Width = 7320
    Me.ScaleHeight = 3135
    Me.ScaleWidth = 7230
End Sub
