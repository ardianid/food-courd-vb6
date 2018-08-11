VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_karyawan_lain 
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   10440
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6015
      Left            =   120
      Negotiate       =   -1  'True
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   5985
      ScaleWidth      =   10185
      TabIndex        =   0
      Top             =   120
      Width           =   10215
      Begin VB.TextBox txt_nama 
         Appearance      =   0  'Flat
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
         Left            =   1920
         TabIndex        =   1
         Top             =   240
         Width           =   5055
      End
      Begin VB.TextBox txt_tempat_lhr 
         Appearance      =   0  'Flat
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
         Left            =   1920
         TabIndex        =   2
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox txt_alamat 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   1920
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   4
         Top             =   1200
         Width           =   5055
      End
      Begin VB.TextBox txt_telp_rmh 
         Appearance      =   0  'Flat
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
         Left            =   1920
         TabIndex        =   5
         Top             =   3000
         Width           =   2535
      End
      Begin VB.TextBox txt_telp_hp 
         Appearance      =   0  'Flat
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
         Left            =   1920
         TabIndex        =   6
         Top             =   3480
         Width           =   3495
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
         Left            =   8640
         TabIndex        =   8
         Top             =   5280
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         Height          =   3255
         Left            =   7320
         TabIndex        =   11
         Top             =   120
         Width           =   2655
         Begin VB.Image img_foto 
            Height          =   3135
            Left            =   0
            Stretch         =   -1  'True
            Top             =   120
            Width           =   2655
         End
      End
      Begin VB.Frame Frame2 
         Height          =   855
         Left            =   7320
         TabIndex        =   9
         Top             =   3360
         Width           =   2655
         Begin VB.CommandButton cmd_foto 
            Caption         =   "F O T O"
            Height          =   495
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.TextBox txt_gaji 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   1920
         TabIndex        =   7
         Top             =   3960
         Width           =   2535
      End
      Begin MSComDlg.CommonDialog cd 
         Left            =   7920
         Top             =   2520
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSMask.MaskEdBox msk_tgl 
         Height          =   375
         Left            =   5520
         TabIndex        =   3
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
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
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   1620
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tempat Lahir"
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
         TabIndex        =   17
         Top             =   840
         Width           =   1290
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pada Tgl."
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
         Left            =   4560
         TabIndex        =   16
         Top             =   840
         Width           =   930
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
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
         TabIndex        =   15
         Top             =   1320
         Width           =   675
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telp. Rumah"
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
         TabIndex        =   14
         Top             =   3120
         Width           =   1260
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telp. HP"
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
         TabIndex        =   13
         Top             =   3600
         Width           =   840
      End
      Begin VB.Line Line1 
         X1              =   360
         X2              =   9960
         Y1              =   5040
         Y2              =   5040
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gaji Pokok"
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
         TabIndex        =   12
         Top             =   4080
         Width           =   1035
      End
   End
End
Attribute VB_Name = "frm_karyawan_lain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim foto As String

Private Sub cmd_foto_Click()
    
    Dim nama_file As String
    On Error Resume Next
    With cd
        .CancelError = True
        .Filter = "File Gambar|*.jpg;*.gif"
        .ShowOpen
    End With
    nama_file = Mid(cd.FileName, InStrRev(cd.FileName, "\"))
    foto = nama_file
    img_foto.Picture = LoadPicture(path_foto & "\foto" & nama_file)

End Sub

Private Sub cmd_simpan_Click()

On Error GoTo er_s

    Dim sql, sql1 As String
    Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
    
    If txt_nama.Text = "" Then
    Dim k As Integer
        k = MsgBox("Nama Pegawai Hrs Diisi", vbOKOnly + vbInformation, "Pesan")
    Exit Sub
    End If
    
    If msk_tgl.Text = "__/__/____" Then
        MsgBox ("Tgl lahir tidak boleh kosong")
        msk_tgl.SetFocus
        Exit Sub
    End If
    
    If MsgBox("Yakin data yang dimasukkan sudah benar....?", vbYesNo + vbQuestion, "Pesan") = vbNo Then
        Exit Sub
    End If
    
If mdl_karyawan = True Then
    sql = "select nama_karyawan from tbl_karyawan where nama_karyawan='" & Trim(txt_nama.Text) & "'"
    rs.Open sql, cn
        If Not rs.EOF Then
            If MsgBox("Data karyawan dengan nama " & Trim(txt_nama.Text) & " Sudah ada,Yakin akan menyimpannya", vbYesNo + vbInformation, "Pesan") = vbYes Then
                isi_dtbs
            Else
                Exit Sub
            End If
        Else
            isi_dtbs
        End If
    rs.Close
    
    MsgBox ("Data berhasil disimpan")
    mdl_karyawan = True
    kosong_semua
    txt_nama.SetFocus
    
ElseIf mdl_karyawan = False Then
    
    sql = "select id from tbl_karyawan where id=" & id_kar
    rs.Open sql, cn
        If Not rs.EOF Then
            
            Dim cek_t
                cek_t = Trim(msk_tgl.Text)
            
            
            sql1 = "update tbl_karyawan set nama_karyawan='" & Trim(txt_nama.Text) & "',tempat_lhr='" & Trim(txt_tempat_lhr.Text) & "',tgl_lhr='" & cek_t & "',alamat='" & Trim(txt_alamat.Text) & "',telp_rumah='" & Trim(txt_telp_rmh.Text) & "',telp_hp='" & Trim(txt_telp_hp.Text) & "',foto='" & foto & "',gaji_pokok=" & CCur(Trim(txt_gaji.Text)) & " where id=" & id_kar
            rs1.Open sql1, cn
            
            MsgBox ("Data berhasil diedit")
            frm_browse_pegawai.isi_k
            Unload Me
            Exit Sub
            
        Else
            MsgBox ("Data yang akan diedit tidak ditemukan")
        End If
        rs.Close
End If
    Exit Sub
    
er_s:
    
    Dim er
        er = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
End Sub

Private Sub isi_dtbs()
On Error Resume Next

    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim cek_tgl
    
        cek_tgl = Trim(msk_tgl.Text)
   
        sql = "insert into tbl_karyawan (nama_karyawan,tempat_lhr,tgl_lhr,alamat,telp_rumah,telp_hp,foto,gaji_pokok)"
        sql = sql & " values ('" & Trim(txt_nama.Text) & "','" & Trim(txt_tempat_lhr.Text) & "','" & cek_tgl & "','" & Trim(txt_alamat.Text) & "','" & Trim(txt_telp_rmh.Text) & "','" & Trim(txt_telp_hp.Text) & "','" & foto & "'," & CCur(Trim(txt_gaji.Text)) & ")"
        rs.Open sql, cn
        
End Sub

Private Sub kosong_semua()
    txt_nama.Text = ""
    txt_tempat_lhr.Text = ""
    msk_tgl.Text = "__/__/____"
    txt_alamat.Text = ""
    txt_telp_rmh.Text = ""
    txt_telp_hp.Text = ""
    img_foto.Picture = LoadPicture()
    txt_gaji.Text = 0
End Sub

Private Sub Form_Activate()
    txt_nama.SetFocus
End Sub

Private Sub Form_Load()
    If mdl_karyawan = True Then
        kosong_semua
        foto = ""
    ElseIf mdl_karyawan = False Then
        kosong_semua
        isi_berdasarkan
    End If
    
    Me.Left = utama.Width \ 2 - Me.Width \ 2
    Me.Top = utama.Height \ 2 - Me.Height \ 2
    
End Sub

Private Sub isi_berdasarkan()
On Error Resume Next
    Dim sql As String
    Dim rs As New ADODB.Recordset
        sql = "select * from tbl_karyawan where id=" & id_kar
        rs.Open sql, cn
            If Not rs.EOF Then
                txt_nama.Text = rs("nama_karyawan")
                
                If Not IsNull(rs("tempat_lhr")) Then
                    txt_tempat_lhr.Text = rs("tempat_lhr")
                End If
                
                If Not IsNull(rs("tgl_lhr")) Then
                    msk_tgl.Text = rs("tgl_lhr")
                End If
                
                If Not IsNull(rs("alamat")) Then
                    txt_alamat.Text = rs("alamat")
                End If
                
                If Not IsNull(rs("telp_rumah")) Then
                    txt_telp_rmh.Text = rs("telp_rumah")
                End If
                
                If Not IsNull(rs("telp_hp")) Then
                    txt_telp_hp.Text = rs("telp_hp")
                End If
                
                If Not IsNull(rs("foto")) And rs("foto") <> "" Then
                    Set img_foto.Picture = LoadPicture(path_foto & "\foto" & "\" & rs("foto"))
                        foto = "\" & rs("foto")
                Else
                    Set img_foto.Picture = LoadPicture("")
                        foto = ""
                End If
                
                If Not IsNull(rs("gaji_pokok")) Then
                    txt_gaji.Text = rs("gaji_pokok")
                Else
                    txt_gaji.Text = 0
                End If
                
          Else
                MsgBox ("Data yang akan diedit tidak ditemukan")
                frm_browse_pegawai.isi_k
                Unload Me
                Exit Sub
          End If
        rs.Close
                
End Sub

Private Sub txt_alamat_GotFocus()
    txt_alamat.SelStart = 0
    txt_alamat.SelLength = Len(txt_alamat)
End Sub

Private Sub txt_gaji_GotFocus()
    txt_gaji.SelStart = 0
    txt_gaji.SelLength = Len(txt_gaji)
End Sub

Private Sub txt_gaji_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub txt_gaji_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If txt_gaji.Text <> "" Then
        txt_gaji.Text = Format(txt_gaji.Text, "###,###,###")
        txt_gaji.SelStart = Len(txt_gaji.Text)
    End If
End Sub

Private Sub txt_gaji_LostFocus()
    If txt_gaji.Text = "" Then
        txt_gaji.Text = 0
    End If
End Sub

Private Sub txt_nama_GotFocus()
    txt_nama.SelStart = 0
    txt_nama.SelLength = Len(txt_nama)
End Sub
Private Sub txt_telp_hp_GotFocus()
    txt_telp_hp.SelStart = 0
    txt_telp_hp.SelLength = Len(txt_telp_hp)
End Sub

Private Sub txt_telp_rmh_GotFocus()
    txt_telp_rmh.SelStart = 0
    txt_telp_rmh.SelLength = Len(txt_telp_rmh)
End Sub

Private Sub txt_tempat_lhr_GotFocus()
    txt_tempat_lhr.SelStart = 0
    txt_tempat_lhr.SelLength = Len(txt_tempat_lhr)
End Sub
Private Sub msk_tgl_GotFocus()
    msk_tgl.SelStart = 0
    msk_tgl.SelLength = Len(msk_tgl)
End Sub

