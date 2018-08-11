VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{1ABFD380-C196-11D2-B0EA-00A024695830}#1.0#0"; "ticon3d6.ocx"
Begin VB.Form frm_hak_akses 
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8400
   ScaleWidth      =   15225
   WindowState     =   2  'Maximized
   Begin TDBContainer3D6Ctl.TDBContainer3D pic_karyawan 
      Height          =   4575
      Left            =   1080
      TabIndex        =   34
      Top             =   3960
      Visible         =   0   'False
      Width           =   5295
      _Version        =   65536
      _ExtentX        =   9340
      _ExtentY        =   8070
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "frm_hak_akses.frx":0000
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "frm_hak_akses.frx":001C
      Childs          =   "frm_hak_akses.frx":00C8
      Begin VB.Frame Frame5 
         Height          =   135
         Index           =   0
         Left            =   240
         TabIndex        =   36
         Top             =   360
         Width           =   4815
      End
      Begin VB.TextBox txt_cari 
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
         Left            =   2040
         TabIndex        =   35
         Top             =   600
         Width           =   3015
      End
      Begin TrueDBGrid60.TDBGrid grd_karyawan 
         Height          =   3375
         Left            =   240
         OleObjectBlob   =   "frm_hak_akses.frx":00E4
         TabIndex        =   37
         Top             =   1080
         Width           =   4815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PENCARIAN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   240
         TabIndex        =   39
         Top             =   120
         Width           =   1065
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Nama Karyawan"
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
         TabIndex        =   38
         Top             =   600
         Width           =   1635
      End
   End
   Begin VB.PictureBox pic_karyawan1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6015
      Left            =   -5400
      ScaleHeight     =   5985
      ScaleWidth      =   5625
      TabIndex        =   14
      Top             =   7920
      Visible         =   0   'False
      Width           =   5655
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         ScaleHeight     =   345
         ScaleWidth      =   5625
         TabIndex        =   16
         Top             =   0
         Width           =   5655
         Begin VB.CommandButton cmd_x 
            Caption         =   "x"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5280
            TabIndex        =   17
            Top             =   0
            Width           =   375
         End
      End
      Begin VB.TextBox txt_cari1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   5415
      End
      Begin TrueDBGrid60.TDBGrid grd_karyawan1 
         Height          =   4695
         Left            =   120
         OleObjectBlob   =   "frm_hak_akses.frx":27DA
         TabIndex        =   18
         Top             =   1200
         Width           =   5415
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
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
         TabIndex        =   19
         Top             =   480
         Width           =   5415
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      ScaleHeight     =   825
      ScaleWidth      =   14985
      TabIndex        =   0
      Top             =   120
      Width           =   15015
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
         Left            =   13200
         TabIndex        =   3
         Top             =   120
         Width           =   1575
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
         Height          =   390
         Left            =   2160
         TabIndex        =   2
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Karyawan"
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
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1620
      End
   End
   Begin VB.PictureBox pic_tambah 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6855
      Left            =   120
      ScaleHeight     =   6825
      ScaleWidth      =   14985
      TabIndex        =   20
      Top             =   1080
      Visible         =   0   'False
      Width           =   15015
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   3255
         Left            =   5760
         ScaleHeight     =   3225
         ScaleWidth      =   1425
         TabIndex        =   30
         Top             =   3480
         Width           =   1455
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   5760
         ScaleHeight     =   705
         ScaleWidth      =   1425
         TabIndex        =   29
         Top             =   2640
         Width           =   1455
         Begin VB.CommandButton cmd_t 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   31
            Top             =   120
            Width           =   1215
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   2415
         Left            =   5760
         ScaleHeight     =   2385
         ScaleWidth      =   1425
         TabIndex        =   24
         Top             =   120
         Width           =   1455
         Begin VB.CheckBox cek_lap_tambah 
            Caption         =   "Laporan"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   1680
            Width           =   1575
         End
         Begin VB.CheckBox cek_hapus_tambah 
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
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   1200
            Width           =   1695
         End
         Begin VB.CheckBox cek_edit_tambah 
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
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   720
            Width           =   975
         End
         Begin VB.CheckBox cek_input_tambah 
            Caption         =   "Input "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   975
         End
      End
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
         Left            =   13440
         TabIndex        =   23
         Top             =   6240
         Width           =   1455
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
         Left            =   11880
         TabIndex        =   22
         Top             =   6240
         Width           =   1455
      End
      Begin TrueDBGrid60.TDBGrid grd_tambah 
         Height          =   6615
         Left            =   120
         OleObjectBlob   =   "frm_hak_akses.frx":4EC9
         TabIndex        =   21
         Top             =   120
         Width           =   5535
      End
      Begin TrueDBGrid60.TDBGrid grd_benar 
         Height          =   6015
         Left            =   7320
         OleObjectBlob   =   "frm_hak_akses.frx":819D
         TabIndex        =   32
         Top             =   120
         Width           =   7575
      End
   End
   Begin VB.PictureBox pic_daftar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6855
      Left            =   120
      ScaleHeight     =   6825
      ScaleWidth      =   14985
      TabIndex        =   4
      Top             =   1080
      Width           =   15015
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   6360
         ScaleHeight     =   705
         ScaleWidth      =   8385
         TabIndex        =   11
         Top             =   6000
         Width           =   8415
         Begin VB.CommandButton cmd_edit 
            Caption         =   "Edit"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   33
            Top             =   120
            Width           =   1695
         End
         Begin VB.CommandButton cmd_tambah 
            Caption         =   "Tambah"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   5400
            TabIndex        =   13
            Top             =   120
            Width           =   1455
         End
         Begin VB.CommandButton cdm_hapus 
            Caption         =   "Hapus"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6960
            TabIndex        =   12
            Top             =   120
            Width           =   1335
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   120
         ScaleHeight     =   705
         ScaleWidth      =   6105
         TabIndex        =   6
         Top             =   6000
         Width           =   6135
         Begin VB.CheckBox cek_lap 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Laporan"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4560
            TabIndex        =   10
            Top             =   240
            Width           =   1215
         End
         Begin VB.CheckBox cek_hapus 
            BackColor       =   &H00C0C0C0&
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
            Height          =   255
            Left            =   2880
            TabIndex        =   9
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox cek_edit 
            BackColor       =   &H00C0C0C0&
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
            Height          =   255
            Left            =   1560
            TabIndex        =   8
            Top             =   240
            Width           =   975
         End
         Begin VB.CheckBox cek_input 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Input"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   240
            Width           =   975
         End
      End
      Begin TrueDBGrid60.TDBGrid grd_daftar 
         Height          =   5775
         Left            =   120
         OleObjectBlob   =   "frm_hak_akses.frx":C650
         TabIndex        =   5
         Top             =   120
         Width           =   14655
      End
   End
End
Attribute VB_Name = "frm_hak_akses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arr_karyawan As New XArrayDB
Dim arr_daftar As New XArrayDB
Dim arr_tambah As New XArrayDB
Dim arr_benar As New XArrayDB
Dim id_kr, id_pemakai, id_w, id_hak As String
Dim Moving As Boolean
Dim yold, xold As Long
Private Sub kosong_benar()
    arr_benar.ReDim 0, 0, 0, 0
    grd_benar.ReBind
    grd_benar.Refresh
End Sub

Private Sub cdm_hapus_Click()

On Error GoTo er_hapus

If arr_daftar.UpperBound(1) > 0 Then

    Dim sql, sql1, sql2 As String
    Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset, rs2 As New ADODB.Recordset
        
    If MsgBox("Yakin Akan dihapus", vbYesNo + vbQuestion, "Pesan") = vbNo Then
        Exit Sub
    End If
    
    sql = "select id from tbl_hak_user where id=" & id_hak
    rs.Open sql, cn
        If Not rs.EOF Then
            
            sql1 = "delete from tbl_hak_user where id=" & id_hak
            rs1.Open sql1, cn
            
        End If
    rs.Close
    
    sql = "select id from tbl_wewenang_user where id=" & id_w
    rs.Open sql, cn
        If Not rs.EOF Then
            
            sql1 = "delete from tbl_wewenang_user where id=" & id_w
            rs1.Open sql1, cn
            
        End If
   rs.Close
    
   Cmd_Tampil_Click
   Exit Sub
    
End If
Exit Sub
    
er_hapus:
    
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
    
End Sub

Private Sub cmd_cancel_Click()
    Cmd_Tampil_Click
    pic_tambah.Visible = False
    Picture1.Enabled = True
End Sub

Private Sub cmd_edit_Click()

On Error GoTo er_handler
    
    If arr_daftar.UpperBound(1) = 0 Then
        Exit Sub
    End If
    
    If cmd_edit.Caption = "Edit" Then
        
        Picture3.Enabled = True
        
     If cek_input.Enabled = False And cek_edit.Enabled = False And cek_hapus.Enabled = False And cek_lap.Enabled = False Then
        
        MsgBox ("Data tidak diperbolehkan diedit")
        Picture3.Enabled = False
        Exit Sub
     End If
     
        cmd_edit.Caption = "Simpan Edit"
        Exit Sub
    End If
    
    If cmd_edit.Caption = "Simpan Edit" Then
        
        Dim sql As String
        Dim rs As New ADODB.Recordset
        Dim tam, ed, hap, lap
        
        If cek_input.Enabled = True Then
            If cek_input.Value = vbChecked Then
                tam = 1
            Else
                tam = 0
            End If
        Else
            tam = 0
        End If
        
        If cek_edit.Enabled = True Then
            If cek_edit.Value = vbChecked Then
                ed = 1
            Else
                ed = 0
            End If
        Else
            ed = 0
        End If
        
        If cek_hapus.Enabled = True Then
            If cek_hapus.Value = vbChecked Then
                hap = 1
            Else
                hap = 0
            End If
        Else
            hap = 0
        End If
        
        If cek_lap.Enabled = True Then
            If cek_lap.Value = vbChecked Then
                lap = 1
            Else
                lap = 0
            End If
        Else
            lap = 0
        End If
        
        sql = "update tbl_hak_user set tambah=" & tam & ",edit=" & ed & ",hapus=" & hap & ",lap=" & lap & " where id=" & id_hak
        rs.Open sql, cn
        
        MsgBox ("Data berhasil disimpan")
        Cmd_Tampil_Click
        Picture3.Enabled = False
        cmd_edit.Caption = "Edit"
        Exit Sub
    End If
    Exit Sub
    
er_handler:
        Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
End Sub

Private Sub Cmd_Simpan_Click()

On Error GoTo er_simpan

    Dim sql, sql1 As String
    Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
    Dim a As Long
    Dim cek_pakai, inp, edt, hps, lap
            
        cn.BeginTrans
            
        For a = 1 To arr_benar.UpperBound(1)
           
           
                    
                    sql = "insert into tbl_wewenang_user (id_user,id_form) values(" & id_pemakai & "," & arr_benar(a, 0) & ")"
                    rs.Open sql, cn
                    
                    sql1 = "select id from tbl_wewenang_user where id_user=" & id_pemakai & " and id_form=" & arr_benar(a, 0)
                    rs1.Open sql1, cn
                        If Not rs1.EOF Then
                            inp = 0
                            If arr_benar(a, 2) <> 0 Then
                                inp = 1
                            End If
                            
                            edt = 0
                            If arr_benar(a, 3) <> 0 Then
                                edt = 1
                            End If
                            
                            hps = 0
                            If arr_benar(a, 4) <> 0 Then
                                hps = 1
                            End If
                            
                            lap = 0
                            If arr_benar(a, 5) <> 0 Then
                                lap = 1
                            End If
                            
                            sql = "insert into tbl_hak_user (id_wewenang,tambah,edit,hapus,lap) values(" & rs1!id & "," & inp & "," & edt & "," & hps & "," & lap & ")"
                            rs.Open sql, cn
                            
                        Else
                            
                            MsgBox ("Data gagal disimpan silahkan coba lagi")
                            cn.RollbackTrans
                            Exit Sub
                            
                        End If
                   rs1.Close
                        
           
        Next a
        MsgBox ("Data berhasil disimpan")
        cn.CommitTrans
        cmd_cancel_Click
        Exit Sub
        
er_simpan:
        cn.RollbackTrans
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear

End Sub

Private Sub cmd_t_Click()

On Error GoTo er_t

    Dim jml_brs
            
            If arr_tambah.UpperBound(1) = 0 Then
                Exit Sub
            End If
            
            If arr_tambah(grd_tambah.Bookmark, 1) = vbChecked Then
                MsgBox ("Tidak bisa ditambahkan karena sudah ada")
                Exit Sub
            End If
            
        arr_benar.ReDim 1, arr_benar.UpperBound(1) + 1, 0, 7
        grd_benar.ReBind
        grd_benar.Refresh
            
            jml_brs = arr_benar.UpperBound(1)
            
            arr_benar(jml_brs, 0) = arr_tambah(grd_tambah.Bookmark, 0)
            arr_benar(jml_brs, 1) = arr_tambah(grd_tambah.Bookmark, 2)
            
            If cek_input_tambah.Value = vbChecked Then
                arr_benar(jml_brs, 2) = vbChecked
            Else
                arr_benar(jml_brs, 2) = vbUnchecked
            End If
            
            If cek_edit_tambah.Value = vbChecked Then
                arr_benar(jml_brs, 3) = vbChecked
            Else
                arr_benar(jml_brs, 3) = vbUnchecked
            End If
            
            If cek_hapus_tambah.Value = vbChecked Then
                arr_benar(jml_brs, 4) = vbChecked
            Else
                arr_benar(jml_brs, 4) = vbUnchecked
            End If
            
            If cek_lap_tambah.Value = vbChecked Then
                arr_benar(jml_brs, 5) = vbChecked
            Else
                arr_benar(jml_brs, 5) = vbUnchecked
            End If
            
        grd_benar.ReBind
        grd_benar.Refresh
            
            arr_tambah(grd_tambah.Bookmark, 1) = vbChecked
            
        grd_tambah.ReBind
        grd_tambah.Refresh
            
        Exit Sub
        
er_t:
        Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
            
End Sub

Private Sub cmd_tambah_Click()
On Error Resume Next
    If txt_nama.Text = "" Then
        MsgBox ("Nama Karyawan harus diisi")
        txt_nama.SetFocus
        Exit Sub
    End If
    
    Dim sql, sql1 As String
    Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
    Dim a As Long
    Dim id_w, nm_form As String
        
        kosong_tambah
        kosong_benar
        
        If id_pemakai = "" Then
            Exit Sub
        End If
        
        sql = "select * from tbl_form"
        rs.Open sql, cn, adOpenKeyset
            If Not rs.EOF Then
                
                rs.MoveLast
                rs.MoveFirst
                
            a = 1
                Do While Not rs.EOF
                        
                        id_w = rs!id
                        nm_form = rs!nama_form
                
                sql1 = "select id_form  from tbl_wewenang_user where id_form=" & rs!id & " and id_user=" & id_pemakai
                rs1.Open sql1, cn
                  If rs1.EOF Then
                    
                    arr_tambah.ReDim 1, a, 0, 3
                    grd_tambah.ReBind
                    grd_tambah.Refresh
                    
                    arr_tambah(a, 0) = id_w
                    arr_tambah(a, 1) = vbUnchecked
                    arr_tambah(a, 2) = nm_form
                    
                    a = a + 1
                  End If
                rs1.Close
                rs.MoveNext
                Loop
                grd_tambah.ReBind
                grd_tambah.Refresh
           End If
        rs.Close
        grd_tambah.MoveFirst
        jangan_cek
        pic_tambah.Visible = True
        Picture1.Enabled = False
        
End Sub

Private Sub jangan_cek()
    cek_input_tambah.Value = vbUnchecked
    cek_edit_tambah.Value = vbUnchecked
    cek_hapus_tambah.Value = vbUnchecked
    cek_lap_tambah.Value = vbUnchecked
End Sub

Private Sub Cmd_Tampil_Click()

On Error GoTo er_tampil

    Dim sql, sql1 As String
    Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
        
        If txt_nama.Text = "" Then
            MsgBox ("Nama karyawan hrs diisi")
            txt_nama.SetFocus
            Exit Sub
        End If
        
        sql = "select id_karyawan,id from qr_user where nama_karyawan ='" & Trim(txt_nama.Text) & "'"
        rs.Open sql, cn
            If Not rs.EOF Then
                
                id_kr = rs("id_karyawan")
                id_pemakai = rs("id")
                
            Else
                
                MsgBox ("Nama karyawan yang dimasukkan tidak ditemukan dalam data pemakai program")
                txt_nama.SetFocus
                Exit Sub
            End If
        rs.Close
        
        kosong_daftar
        Dim a As Long
        Dim nama_form, id_dia As String
        
        sql = "select id_wewenang,nama_form from qr_wewenang where id_user=" & id_pemakai & " order by nama_form"
        rs.Open sql, cn, adOpenKeyset
            If Not rs.EOF Then
                
                rs.MoveLast
                rs.MoveFirst
            
            a = 1
                Do While Not rs.EOF
                    arr_daftar.ReDim 1, a, 0, 3
                    grd_daftar.ReBind
                    grd_daftar.Refresh
                        
                        id_dia = rs!id_wewenang
                        If Not IsNull(rs!nama_form) Then
                            nama_form = rs!nama_form
                        Else
                            nama_form = ""
                        End If
                        
                   arr_daftar(a, 0) = id_dia
                   arr_daftar(a, 1) = a
                   arr_daftar(a, 2) = nama_form
                   
               a = a + 1
               rs.MoveNext
               Loop
               grd_daftar.ReBind
               grd_daftar.Refresh
            Else
                MsgBox ("Belum ada form yang boleh diakses oleh user")
                Exit Sub
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
    pic_karyawan.Visible = False
    txt_nama.SetFocus
End Sub

Private Sub Form_Activate()
    txt_nama.SetFocus
End Sub

Private Sub Form_Load()
    
    grd_karyawan.Array = arr_karyawan
    
    isi_karyawan
    
    grd_daftar.Array = arr_daftar
    
    grd_tambah.Array = arr_tambah
    
    grd_benar.Array = arr_benar
    
    kosong_daftar
    
    kosong_benar
    
    
    
    Me.Height = 8115
    Me.Width = 12030
    Me.ScaleHeight = 7605
    Me.ScaleWidth = 11910
    
End Sub

Private Sub kosong_tambah()
    arr_tambah.ReDim 0, 0, 0, 0
    grd_tambah.ReBind
    grd_tambah.Refresh
End Sub

Private Sub kosong_karyawan()
    arr_karyawan.ReDim 0, 0, 0, 0
    grd_karyawan.ReBind
    grd_karyawan.Refresh
End Sub

Private Sub kosong_daftar()
    arr_daftar.ReDim 0, 0, 0, 0
    grd_daftar.ReBind
    grd_daftar.Refresh
End Sub

Private Sub isi_karyawan()

On Error GoTo er_k

    Dim sql As String
    Dim rs_karyawan As New ADODB.Recordset
        
        kosong_karyawan
        
        sql = "select id_karyawan,nama_karyawan from qr_user order by nama_karyawan"
        rs_karyawan.Open sql, cn, adOpenKeyset
            If Not rs_karyawan.EOF Then
                
                rs_karyawan.MoveLast
                rs_karyawan.MoveFirst
                
                lanjut_karyawan rs_karyawan
                
            End If
       rs_karyawan.Close
        
       Exit Sub
       
er_k:
       Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
        
End Sub

Private Sub lanjut_karyawan(rs_karyawan As Recordset)
    Dim id_k, nm As String
    Dim a  As Long
        
        a = 1
            Do While Not rs_karyawan.EOF
                arr_karyawan.ReDim 1, a, 0, 2
                grd_karyawan.ReBind
                grd_karyawan.Refresh
                    
                    id_k = rs_karyawan("id_karyawan")
                    nm = rs_karyawan("nama_karyawan")
                    
               arr_karyawan(a, 0) = id_k
               arr_karyawan(a, 1) = nm
           a = a + 1
           rs_karyawan.MoveNext
           Loop
           grd_karyawan.ReBind
           grd_karyawan.Refresh
                
End Sub

Private Sub grd_benar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        
        If arr_benar.UpperBound(1) = 0 Then
            Exit Sub
        End If
            
        Dim a
            For a = 1 To arr_tambah.UpperBound(1)
                If arr_tambah(a, 0) = arr_benar(grd_benar.Bookmark, 0) Then
                    arr_tambah(a, 1) = vbUnchecked
                End If
            Next a
        
        grd_tambah.ReBind
        grd_tambah.Refresh
        
        If arr_benar.UpperBound(1) > 1 Then
            grd_benar.Delete
        Else
            arr_benar.ReDim 0, 0, 0, 0
        End If
        
        grd_benar.ReBind
        grd_benar.Refresh
        
    End If
End Sub

Private Sub grd_daftar_Click()
On Error Resume Next
    If arr_daftar.UpperBound(1) > 0 Then
        
        id_w = arr_daftar(grd_daftar.Bookmark, 0)
        
        Dim sql As String
        Dim rs As New ADODB.Recordset
        
        cek_input.Value = vbUnchecked
        cek_edit.Value = vbUnchecked
        cek_lap.Value = vbUnchecked
        cek_hapus.Value = vbUnchecked
        
        cek_input.Enabled = True
        cek_edit.Enabled = True
        cek_hapus.Enabled = True
        cek_lap.Enabled = True
        
        If arr_daftar(grd_daftar.Bookmark, 2) = "Form Stock" Then
            
            cek_edit.Enabled = False
            cek_hapus.Enabled = False
            
        End If
    
        If arr_daftar(grd_daftar.Bookmark, 2) = "Form Data Jam Kerja" Then
            
            cek_lap.Enabled = False
            
        End If
        
        If arr_daftar(grd_daftar.Bookmark, 2) = "Form Pembagian Tugas Kerja" Then
            
            cek_edit.Enabled = False
            cek_lap.Enabled = False
            
        End If
        
        If arr_daftar(grd_daftar.Bookmark, 2) = "Form Pembatalan Transaksi" Then
            
            cek_edit.Enabled = False
            cek_hapus.Enabled = False
        
        End If
        
        If arr_daftar(grd_daftar.Bookmark, 2) = "Form Transaksi Billing" Then
            
            cek_edit.Enabled = False
            
        End If
        
        If arr_daftar(grd_daftar.Bookmark, 2) = "Form Data Member" Then
            
            cek_lap.Enabled = False
            
        End If
        
        
        If arr_daftar(grd_daftar.Bookmark, 2) = "Form Penyesuaian Stock" Or arr_daftar(grd_daftar.Bookmark, 2) = "Form Historical Stock" _
           Or arr_daftar(grd_daftar.Bookmark, 2) = "Form Inventory" Or arr_daftar(grd_daftar.Bookmark, 2) = "Form Penyesuaian Inventory" _
           Or arr_daftar(grd_daftar.Bookmark, 2) = "Form Transakai Penjualan" Or arr_daftar(grd_daftar.Bookmark, 2) = "Form Laporan Penjualan" _
           Or arr_daftar(grd_daftar.Bookmark, 2) = "Form Password" Or arr_daftar(grd_daftar.Bookmark, 2) = "Form Ganti Password" Or arr_daftar(grd_daftar.Bookmark, 2) = "Form Browse Penjualan" _
           Or arr_daftar(grd_daftar.Bookmark, 2) = "Form Data Biaya-Biaya" Or arr_daftar(grd_daftar.Bookmark, 2) = "Form Penggajian" Or arr_daftar(grd_daftar.Bookmark, 2) = "Form Backup Database" Then
            
           
            cek_input.Enabled = False
            cek_edit.Enabled = False
            cek_hapus.Enabled = False
            cek_lap.Enabled = False
        
        End If
        
        sql = "select * from qr_hak where id_wewenang=" & Trim(arr_daftar(grd_daftar.Bookmark, 0))
        rs.Open sql, cn
            If Not rs.EOF Then
                
                id_hak = rs!id_hak
                
                If rs!tambah = 1 Then
                    cek_input.Value = vbChecked
                End If
                
                If rs!edit = 1 Then
                    cek_edit.Value = vbChecked
                End If
                
                If rs!hapus = 1 Then
                    cek_hapus.Value = vbChecked
                End If
                
                If rs!lap = 1 Then
                    cek_lap.Value = vbChecked
                End If
                
            End If
        rs.Close
     End If
End Sub

Private Sub grd_daftar_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    grd_daftar_Click
End Sub

Private Sub grd_karyawan_Click()
    On Error Resume Next
        If arr_karyawan.UpperBound(1) > 0 Then
            id_kr = arr_karyawan(grd_karyawan.Bookmark, 0)
        End If
End Sub

Private Sub grd_karyawan_DblClick()
    If arr_karyawan.UpperBound(1) > 0 Then
        txt_nama.Text = arr_karyawan(grd_karyawan.Bookmark, 1)
        pic_karyawan.Visible = False
        txt_nama.SetFocus
    End If
End Sub

Private Sub grd_karyawan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        grd_karyawan_DblClick
    End If
    
    If KeyCode = vbKeyEscape Then
        pic_karyawan.Visible = False
        txt_nama.SetFocus
    End If
    
End Sub

Private Sub grd_karyawan_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    grd_karyawan_Click
End Sub

Private Sub grd_tambah_Click()

On Error GoTo er_t

    If arr_tambah.UpperBound(1) > 0 Then
        
        cek_input_tambah.Enabled = True
        cek_edit_tambah.Enabled = True
        cek_hapus_tambah.Enabled = True
        cek_lap_tambah.Enabled = True
        
        cek_input_tambah.Value = vbChecked
        cek_edit_tambah.Value = vbChecked
        cek_hapus_tambah.Value = vbChecked
        cek_lap_tambah.Value = vbChecked
        
        If arr_tambah(grd_tambah.Bookmark, 2) = "Form Stock" Then
            
            cek_edit_tambah.Value = vbUnchecked
            cek_hapus_tambah.Value = vbUnchecked
            
            cek_edit_tambah.Enabled = False
            cek_hapus_tambah.Enabled = False
            
        End If
    
        If arr_tambah(grd_tambah.Bookmark, 2) = "Form Data Jam Kerja" Then
            
            cek_lap_tambah.Value = vbUnchecked
            
            cek_lap_tambah.Enabled = False
            
        End If

        If arr_tambah(grd_tambah.Bookmark, 2) = "Form Data Member" Then
            
            cek_lap_tambah.Value = vbUnchecked
            
            cek_lap_tambah.Enabled = False
            
        End If


        If arr_tambah(grd_tambah.Bookmark, 2) = "Form Pembagian Tugas Kerja" Then
            
            cek_edit_tambah.Value = vbUnchecked
            cek_lap_tambah.Value = vbUnchecked
            
            cek_edit_tambah.Enabled = False
            cek_lap_tambah.Enabled = False
            
        End If
        
        If arr_tambah(grd_tambah.Bookmark, 2) = "Form Pembatalan Transaksi" Then
            
            cek_edit_tambah.Value = vbUnchecked
            cek_hapus_tambah.Value = vbUnchecked
            
            cek_edit_tambah.Enabled = False
            cek_hapus_tambah.Enabled = False
        
        End If
        
        If arr_tambah(grd_tambah.Bookmark, 2) = "Form Transaksi Billing" Then
            
            cek_edit_tambah.Value = vbUnchecked
            
            cek_edit_tambah.Enabled = False
            
        End If
        
        If arr_tambah(grd_tambah.Bookmark, 2) = "Form Penyesuaian Stock" Or arr_tambah(grd_tambah.Bookmark, 2) = "Form Historical Stock" _
           Or arr_tambah(grd_tambah.Bookmark, 2) = "Form Inventory" Or arr_tambah(grd_tambah.Bookmark, 2) = "Form Penyesuaian Inventory" _
           Or arr_tambah(grd_tambah.Bookmark, 2) = "Form Transakai Penjualan" Or arr_tambah(grd_tambah.Bookmark, 2) = "Form Laporan Penjualan" _
           Or arr_tambah(grd_tambah.Bookmark, 2) = "Form Password" Or arr_tambah(grd_tambah.Bookmark, 2) = "Form Ganti Password" Or arr_tambah(grd_tambah.Bookmark, 2) = "Form Browse Penjualan" _
           Or arr_tambah(grd_tambah.Bookmark, 2) = "Form Data Biaya-Biaya" Or arr_tambah(grd_tambah.Bookmark, 2) = "Form Penggajian" _
           Or arr_tambah(grd_tambah.Bookmark, 2) = "Form Lap Perkasir" Or arr_tambah(grd_tambah.Bookmark, 2) = "Form Backup Database" Or arr_tambah(grd_tambah.Bookmark, 2) = "Form Laporan PerCounter" Or arr_tambah(grd_tambah.Bookmark, 2) = "Form Laporan PerCounter Berdasarkan Disc" _
           Or arr_tambah(grd_tambah.Bookmark, 2) = "Form Slip Pembayaran" Or arr_tambah(grd_tambah.Bookmark, 2) = "Form Persentase" Then
            
            cek_input_tambah.Value = vbUnchecked
            cek_edit_tambah.Value = vbUnchecked
            cek_hapus_tambah.Value = vbUnchecked
            cek_lap_tambah.Value = vbUnchecked
            
            cek_input_tambah.Enabled = False
            cek_edit_tambah.Enabled = False
            cek_hapus_tambah.Enabled = False
            cek_lap_tambah.Enabled = False
        
        End If
    End If
    
    Exit Sub
    
er_t:
    Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
    
End Sub

Private Sub grd_tambah_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    grd_tambah_Click
End Sub

'Private Sub grd_tambah_AfterColUpdate(ByVal ColIndex As Integer)
'    If ColIndex = 1 Then
'        arr_tambah(grd_tambah.Bookmark, ColIndex) = grd_tambah.Columns(ColIndex).Text
'    End If
'
'    If ColIndex = 3 Then
'        arr_tambah(grd_tambah.Bookmark, ColIndex) = grd_tambah.Columns(ColIndex).Text
'    End If
'
'    If ColIndex = 4 Then
'        arr_tambah(grd_tambah.Bookmark, ColIndex) = grd_tambah.Columns(ColIndex).Text
'    End If
'
'    If ColIndex = 5 Then
'        arr_tambah(grd_tambah.Bookmark, ColIndex) = grd_tambah.Columns(ColIndex).Text
'    End If
'
'    If ColIndex = 6 Then
'        arr_tambah(grd_tambah.Bookmark, ColIndex) = grd_tambah.Columns(ColIndex).Text
'    End If
'End Sub

Private Sub pic_karyawan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        pic_karyawan.Visible = False
        txt_nama.SetFocus
    End If
End Sub

Private Sub pic_karyawan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = True
If Moving = True Then
   yold = Y
   xold = X
End If
End Sub

Private Sub pic_karyawan_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Moving = True Then
   pic_karyawan.Top = pic_karyawan.Top - (yold - Y)
   pic_karyawan.Left = pic_karyawan.Left - (xold - X)
End If

End Sub

Private Sub pic_karyawan_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = False
End Sub

Private Sub txt_cari_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        pic_karyawan.Visible = False
        txt_nama.SetFocus
    End If
    
    If KeyCode = 13 Then
        
        txt_nama.Text = arr_karyawan(grd_karyawan.Bookmark, 1)
        pic_karyawan.Visible = False
        txt_nama.SetFocus
    End If
End Sub

Private Sub txt_cari_KeyUp(KeyCode As Integer, Shift As Integer)

On Error GoTo er_cari

    Dim sql As String
    Dim rs_karyawan As New ADODB.Recordset
        
        sql = "select id_karyawan,nama_karyawan from qr_user"
            
            If txt_cari.Text <> "" Then
                sql = sql & " where nama_karyawan like '%" & Trim(txt_cari.Text) & "%'"
            End If
                
       sql = sql & " order by nama_karyawan"
       rs_karyawan.Open sql, cn, adOpenKeyset
        If Not rs_karyawan.EOF Then
            
            rs_karyawan.MoveLast
            rs_karyawan.MoveFirst
            
            lanjut_karyawan rs_karyawan
        End If
      rs_karyawan.Close
        
      Exit Sub
      
er_cari:
      Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
        
End Sub
Private Sub txt_nama_GotFocus()
    txt_nama.SelStart = 0
    txt_nama.SelLength = Len(txt_nama)
End Sub

Private Sub txt_nama_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        kosong_daftar
        txt_nama.Text = ""
        txt_cari.Text = ""
        pic_karyawan.Visible = True
        txt_cari.SetFocus
    End If
End Sub
