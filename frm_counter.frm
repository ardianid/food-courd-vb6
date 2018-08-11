VERSION 5.00
Begin VB.Form frm_counter 
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   7845
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5295
      Left            =   0
      ScaleHeight     =   5265
      ScaleWidth      =   7785
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      Begin VB.Frame Frame2 
         Caption         =   "Data counter"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   120
         TabIndex        =   12
         Top             =   2040
         Width           =   7575
         Begin VB.TextBox txt_counter 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2520
            TabIndex        =   5
            Top             =   960
            Width           =   4095
         End
         Begin VB.TextBox txt_presentasi 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2520
            TabIndex        =   6
            Text            =   "0"
            Top             =   1440
            Width           =   975
         End
         Begin VB.TextBox txt_kode 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2520
            TabIndex        =   4
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   240
            TabIndex        =   16
            Top             =   960
            Width           =   585
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Presentasi"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   240
            TabIndex        =   15
            Top             =   1440
            Width           =   990
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   3600
            TabIndex        =   14
            Top             =   1440
            Width           =   270
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kode"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   240
            TabIndex        =   13
            Top             =   480
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Data Pemilik Counter"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   120
         TabIndex        =   8
         Top             =   0
         Width           =   7575
         Begin VB.TextBox txt_telp 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1800
            TabIndex        =   3
            Top             =   1320
            Width           =   2055
         End
         Begin VB.TextBox txt_alamat 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1800
            TabIndex        =   2
            Top             =   840
            Width           =   5535
         End
         Begin VB.TextBox txt_nama 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1800
            TabIndex        =   1
            Top             =   360
            Width           =   3375
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telp."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   240
            TabIndex        =   11
            Top             =   1440
            Width           =   495
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Alamat"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   240
            TabIndex        =   10
            Top             =   960
            Width           =   675
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Pemilik"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   240
            TabIndex        =   9
            Top             =   480
            Width           =   1290
         End
      End
      Begin VB.CommandButton cmd_simpan 
         Caption         =   "Simpan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6600
         TabIndex        =   7
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   7680
         Y1              =   4320
         Y2              =   4320
      End
   End
End
Attribute VB_Name = "frm_counter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_simpan_Click()

On Error GoTo er_s

    Dim sql, sql1, sql2 As String
    Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset, rs2 As New ADODB.Recordset

        If txt_kode.Text = "" Or txt_counter.Text = "" Or txt_presentasi.Text = "" Or txt_nama.Text = "" Or txt_alamat.Text = "" Or txt_telp.Text = "" Then
            MsgBox ("Semua data harus diisi")
            Exit Sub
        End If
        
If mdl_counter = True Then
        
    sql = "select kode from tbl_counter where kode='" & Trim(txt_kode.Text) & "'"
    rs.Open sql, cn
        If Not rs.EOF Then
            If MsgBox("Kode Counter " & txt_kode.Text & " sudah ada", vbOKOnly + vbQuestion, "Pesan") = vbOK Then
                txt_kode.SetFocus
                Exit Sub
            End If
        Else
            isi
        End If
    rs.Close
    
    Exit Sub
    
ElseIf mdl_counter = False Then
    
    If MsgBox("Yakin data yang dimasukkan sudah benar....?", vbYesNo + vbQuestion, "Pesan") = vbNo Then
        Exit Sub
    End If
    
    cn.BeginTrans
    
    sql = "select id from tbl_counter where id=" & id_counter
    rs.Open sql, cn
        If Not rs.EOF Then
            Dim presentasi
                presentasi = Trim(txt_presentasi.Text) & "%"
            sql1 = "update tbl_counter set kode='" & Trim(txt_kode.Text) & "', nama_counter='" & Trim(txt_counter.Text) & "',presentasi_p='" & presentasi & "',nama_pemilik='" & Trim(txt_nama.Text) & "',alamat='" & Trim(txt_alamat.Text) & "',telp='" & Trim(txt_telp.Text) & "' where id=" & id_counter
            rs1.Open sql1, cn
            
            sql2 = "select count(kode) as jml_kode from tbl_counter where kode='" & Trim(txt_kode.Text) & "'"
            rs2.Open sql2, cn
                If Not rs2.EOF Then
                    Dim jml As Double
                        jml = rs2("jml_kode")
                            If jml > 1 Then
                                MsgBox ("Kode yang anda masukkan sudah ada")
                                cn.RollbackTrans
                                txt_kode.SetFocus
                                Exit Sub
                            Else
                                MsgBox ("Data berhasil diedit")
                                cn.CommitTrans
                            End If
                Else
                    MsgBox ("Data berhasil diedit")
                    cn.CommitTrans
                End If
            rs2.Close
        Else
            
            MsgBox ("Data yang akan diedit tidak ditemukan")
            
        End If
   rs.Close
   frm_browse_counter.isi_counter
   Unload Me
   Exit Sub
   
End If
    
Exit Sub

er_s:

    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
        
End Sub

Private Sub isi()
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim presentasi
        
        
        
        presentasi = Trim(txt_presentasi.Text) & "%"
        
        sql = "insert into tbl_counter (kode,nama_counter,presentasi_p,nama_pemilik,alamat,telp)"
        sql = sql & " values ('" & Trim(txt_kode.Text) & "', '" & Trim(txt_counter.Text) & "','" & presentasi & "','" & Trim(txt_nama.Text) & "','" & Trim(txt_alamat.Text) & "','" & Trim(txt_telp.Text) & "')"
        rs.Open sql, cn
        
        MsgBox ("Data berhasil disimpan")
        mdl_counter = True
        kosong
        txt_nama.SetFocus
        
End Sub

Private Sub kosong()
    txt_nama.Text = ""
    txt_alamat.Text = ""
    txt_telp.Text = ""
    txt_counter.Text = ""
    txt_kode.Text = ""
  '  txt_presentasi.Text = ""
End Sub

Private Sub Form_Activate()
    txt_nama.SetFocus
'     txt_kode.SetFocus
End Sub

Private Sub Form_Load()
    
    If mdl_counter = True Then
        kosong
    Else
        jangan_kosong
    End If
    
    Me.Left = utama.Width \ 2 - Me.Width \ 2
    Me.Top = utama.Height \ 2 - Me.Height \ 2 - 190
    
End Sub

Private Sub jangan_kosong()

On Error GoTo jangan


    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim pres
        
        sql = "select * from tbl_counter where id=" & id_counter
        rs.Open sql, cn
            If Not rs.EOF Then
                
                If Not IsNull(rs("nama_pemilik")) Then
                    txt_nama.Text = rs("nama_pemilik")
                Else
                    txt_nama.Text = ""
                End If
                
                If Not IsNull(rs("alamat")) Then
                    txt_alamat.Text = rs("alamat")
                Else
                    txt_alamat.Text = ""
                End If
                
                If Not IsNull(rs("telp")) Then
                    txt_telp.Text = rs("telp")
                Else
                    txt_telp.Text = ""
                End If
                
                If Not IsNull(rs("kode")) Then
                    txt_kode.Text = rs("kode")
                Else
                    txt_kode.Text = ""
                End If
                
                If Not IsNull(rs("nama_counter")) Then
                    txt_counter.Text = rs("nama_counter")
                Else
                    txt_counter = ""
                End If
            
                If Not IsNull(rs("presentasi_p")) Then
                    pres = Trim(rs("presentasi_p"))
                    pres = Len(pres)
                    pres = Mid(Trim(rs("presentasi_p")), 1, pres - 1)
                    txt_presentasi.Text = pres
                Else
                    txt_presentasi.Text = ""
                End If
            Else
                MsgBox ("Data yang akan diedit tidak ditemukan")
                frm_browse_counter.isi_counter
                Unload Me
                Exit Sub
            End If
       rs.Close
    
    Exit Sub
    
jangan:
        Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
    
End Sub

Private Sub txt_alamat_GotFocus()
    txt_alamat.SelStart = 0
    txt_alamat.SelLength = Len(txt_alamat)
End Sub

Private Sub txt_counter_GotFocus()
    txt_counter.SelStart = 0
    txt_counter.SelLength = Len(txt_counter)
End Sub

Private Sub txt_nama_GotFocus()
    txt_nama.SelStart = 0
    txt_nama.SelLength = Len(txt_nama)
End Sub

Private Sub txt_presentasi_GotFocus()
    txt_presentasi.SelStart = 0
    txt_presentasi.SelLength = Len(txt_presentasi)
End Sub

Private Sub txt_presentasi_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = Asc(",")) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub txt_telp_GotFocus()
    txt_telp.SelStart = 0
    txt_telp.SelLength = Len(txt_telp)
End Sub
