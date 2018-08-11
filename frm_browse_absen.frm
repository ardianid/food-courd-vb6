VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_browse_absen 
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8400
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      ScaleHeight     =   705
      ScaleWidth      =   14865
      TabIndex        =   13
      Top             =   7320
      Width           =   14895
      Begin VB.CommandButton cmd_export 
         Caption         =   "Export"
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
         Left            =   3240
         TabIndex        =   18
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmd_cetak 
         Caption         =   "Cetak"
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
         Left            =   1680
         TabIndex        =   17
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmd_setup 
         Caption         =   "Page Setup"
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
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   1455
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
         Left            =   13320
         TabIndex        =   15
         Top             =   120
         Width           =   1455
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
         Left            =   11760
         TabIndex        =   14
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4815
      Left            =   120
      ScaleHeight     =   4785
      ScaleWidth      =   14865
      TabIndex        =   11
      Top             =   2400
      Width           =   14895
      Begin MSComDlg.CommonDialog cd 
         Left            =   1560
         Top             =   2400
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin TrueDBGrid60.TDBGrid grd_absen 
         Height          =   4575
         Left            =   120
         OleObjectBlob   =   "frm_browse_absen.frx":0000
         TabIndex        =   12
         Top             =   120
         Width           =   14655
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   120
      ScaleHeight     =   2145
      ScaleWidth      =   14865
      TabIndex        =   0
      Top             =   120
      Width           =   14895
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
         Left            =   13320
         TabIndex        =   5
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox txt_nama 
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
         Left            =   2400
         TabIndex        =   4
         Top             =   1680
         Width           =   3735
      End
      Begin MSMask.MaskEdBox msk_tgl 
         Height          =   375
         Left            =   2400
         TabIndex        =   1
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
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
      Begin MSMask.MaskEdBox msk_tgl1 
         Height          =   375
         Left            =   4320
         TabIndex        =   2
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
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
      Begin MSMask.MaskEdBox msk_jam 
         Height          =   375
         Left            =   2400
         TabIndex        =   3
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
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
      Begin VB.Label Label5 
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
         Left            =   480
         TabIndex        =   10
         Top             =   1680
         Width           =   1620
      End
      Begin VB.Label Label4 
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
         TabIndex        =   9
         Top             =   1200
         Width           =   1110
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "S/d"
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
         Left            =   3840
         TabIndex        =   8
         Top             =   720
         Width           =   330
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl."
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
         TabIndex        =   7
         Top             =   720
         Width           =   360
      End
      Begin VB.Line Line1 
         X1              =   360
         X2              =   14760
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pencarian"
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
         Left            =   240
         TabIndex        =   6
         Top             =   120
         Width           =   1140
      End
   End
End
Attribute VB_Name = "frm_browse_absen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arr_absen As New XArrayDB
Dim id_absen As String

Private Sub cmd_cetak_Click()
On Error GoTo er_printer

    With grd_absen.PrintInfo
        
        .PageHeaderFont.Bold = True
        .PageHeaderFont.Italic = True
        .PageHeaderFont.Size = 10
        .PageHeader = "Laporan Data Absen Karyawan"
        .RepeatColumnHeaders = True
        .PageFooter = "\tPage: \p" & "..." & id_user
        .PrintPreview
    End With
    Exit Sub
    
er_printer:
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
End Sub

Private Sub cmd_edit_Click()
    If cmd_edit.Caption = "Edit" Then
    
        cmd_edit.Caption = "Read Only"
        grd_absen.Columns(7).Locked = False
        grd_absen.Columns(8).Locked = False
        grd_absen.MoveFirst
        Exit Sub
        
    Else
            
        cmd_edit.Caption = "Edit"
        grd_absen.Columns(7).Locked = True
        grd_absen.Columns(8).Locked = True
        cmd_tampil_Click
        
    End If
        
End Sub

Private Sub cmd_export_Click()
    On Error Resume Next

    cd.ShowSave
    grd_absen.ExportToFile cd.FileName, False
    
End Sub

Private Sub cmd_hapus_Click()
    On Error GoTo er_h
    
    If arr_absen.UpperBound(1) > 0 Then
        Dim sql, sql1 As String
        Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
            
            If MsgBox("Yakin akan dihapus.........?", vbYesNo + vbQuestion, "Pesan") = vbNo Then
                Exit Sub
            End If
            
            sql = "select id from tr_absen where id=" & id_absen
            rs.Open sql, cn
                If Not rs.EOF Then
                    sql1 = "delete from tr_absen where id=" & id_absen
                    rs1.Open sql1, cn
                Else
                    MsgBox ("Data yang akan dihapus tidak ditemukan")
                End If
            rs.Close
            cmd_tampil_Click
            Exit Sub
    End If
    Exit Sub
    
er_h:
        Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
End Sub

Private Sub cmd_setup_Click()
On Error GoTo er_set
    With grd_absen.PrintInfo
        .PageSetup
    End With
    Exit Sub
er_set:
        Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
End Sub

Private Sub cmd_tampil_Click()
    Call isi
End Sub

Private Sub Form_Load()

    grd_absen.Array = arr_absen
    
    kosong_absen
    
    Call cari_wewenang("Form Data Absensi Karyawan")
      
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
        
        If lap_form = True Then
            cmd_setup.Enabled = True
            cmd_cetak.Enabled = True
            cmd_export.Enabled = True
        Else
            cmd_setup.Enabled = False
            cmd_cetak.Enabled = False
            cmd_export.Enabled = False
        End If
    
End Sub

Private Sub kosong_absen()
    arr_absen.ReDim 0, 0, 0, 0
    grd_absen.ReBind
    grd_absen.Refresh
End Sub

Private Sub isi()

On Error GoTo er_isi

    Dim sql As String
    Dim rs As New ADODB.Recordset
    
        
        kosong_absen
        
        
                grd_absen.Columns(4).FooterText = "Total Data : " & 0
                grd_absen.Columns(5).FooterText = "Hadir : " & 0
                grd_absen.Columns(6).FooterText = "Alpha : " & 0
                grd_absen.Columns(7).FooterText = "Izin  : " & 0
                grd_absen.Columns(8).FooterText = "Sakit : " & 0
    
        sql = "select * from qr_absen "
            
        If msk_tgl.Text <> "__/__/____" Or msk_tgl1.Text <> "__/__/____" Or msk_jam.Text <> "__:__:__" Or txt_nama.Text <> "" Then
                
            sql = sql & " where"
                
                If msk_tgl.Text <> "__/__/____" And msk_tgl1.Text = "__/__/____" Then
                    sql = sql & " tgl = datevalue('" & Trim(msk_tgl.Text) & "')"
                End If
                
                If msk_tgl1.Text <> "__/__/____" And msk_tgl.Text = "__/__/____" Then
                    sql = sql & " tgl = datevalue('" & Trim(msk_tgl1.Text) & "')"
                End If
                
                If msk_tgl.Text <> "__/__/____" And msk_tgl1.Text <> "__/__/____" Then
                    sql = sql & " tgl >= datevalue('" & Trim(msk_tgl.Text) & "') and tgl <= datevalue('" & Trim(msk_tgl1.Text) & "')"
                End If
                
                If msk_jam.Text <> "__:__:__" And msk_tgl.Text = "__/__/____" And msk_tgl1.Text = "__/__/____" Then
                    sql = sql & " jam_masuk= timevalue('" & Trim(msk_jam.Text) & "')"
                End If
                
                If msk_jam.Text <> "__:__:__" And (msk_tgl.Text <> "__/__/____" Or msk_tgl1.Text <> "__/__/____") Then
                    sql = sql & " and jam_masuk= timevalue('" & Trim(msk_jam.Text) & "')"
                End If
                
                If txt_nama.Text <> "" And msk_jam.Text = "__:__:__" And msk_tgl.Text = "__/__/____" And msk_tgl1.Text = "__/__/____" Then
                    sql = sql & " nama_karyawan like '%" & Trim(txt_nama.Text) & "%'"
                End If
                
                If txt_nama.Text <> "" And (msk_jam.Text <> "__:__:__" Or msk_tgl.Text <> "__/__/____" Or msk_tgl1.Text <> "__/__/____") Then
                    sql = sql & " and nama_karyawan like '%" & Trim(txt_nama.Text) & "%'"
                End If
        End If
            
        sql = sql & " order by tgl,jam"
        rs.Open sql, cn, adOpenKeyset
            If Not rs.EOF Then
                
                rs.MoveLast
                rs.MoveFirst
                
                lanjut_ah rs
            End If
        rs.Close
        
        Exit Sub
        
er_isi:
        Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
End Sub

Private Sub lanjut_ah(rs As Recordset)
    Dim id_a, tgl, jam, jam_masuk, nama, masuk, ket, user As String
    Dim a, b As Long
    Dim hadir, alpha, izin, sakit As Double
            
        a = 1
        b = 1
        hadir = 0
        alpha = 0
        izin = 0
        sakit = 0
            Do While Not rs.EOF
                arr_absen.ReDim 1, a, 0, 10
                grd_absen.ReBind
                grd_absen.Refresh
                    
                    id_a = rs("id")
                    
                    If Not IsNull(rs("tgl")) Then
                        tgl = rs("tgl")
                    Else
                        tgl = ""
                    End If
                    
                    If Not IsNull(rs("jam")) Then
                        jam = Format(rs("jam"), "hh:mm:ss")
                    Else
                        jam = ""
                    End If
                    
                    If Not IsNull(rs("jam_masuk")) Then
                        jam_masuk = Format(rs("jam_masuk"), "hh:mm:ss")
                    Else
                        jam_masuk = ""
                    End If
                    
                    If Not IsNull(rs("nama_karyawan")) Then
                        nama = rs("nama_karyawan")
                    Else
                        nama = ""
                    End If
                    
                    If Not IsNull(rs("masuk_gak")) Then
                        
                        If rs("masuk_gak") = 1 Then
                            hadir = hadir + 1
                        End If
                        
                        If rs("masuk_gak") = 2 Then
                            alpha = alpha + 1
                        End If
                        
                        If rs("masuk_gak") = 3 Then
                            izin = izin + 1
                        End If
                        
                        If rs("masuk_gak") = 4 Then
                            sakit = sakit + 1
                        End If
                        
                        masuk = absen_sebenarnya(rs("masuk_gak"))
                        
                    Else
                        masuk = ""
                    End If
                    
                    If Not IsNull(rs("ket")) Then
                       ket = rs("ket")
                    Else
                        ket = "-"
                    End If
                    
                    If Not IsNull(rs("nama_user")) Then
                        user = rs("nama_user")
                    Else
                        user = ""
                    End If
                    
                If a > 1 Then
                    If tgl <> arr_absen(a - 1, 2) Then
                        b = b + 1
                    End If
                End If
                        
                arr_absen(a, 0) = id_a
                arr_absen(a, 1) = b
                arr_absen(a, 2) = tgl
                arr_absen(a, 3) = jam
                arr_absen(a, 4) = jam_masuk
                arr_absen(a, 5) = nama
                arr_absen(a, 6) = user
                arr_absen(a, 7) = masuk
                arr_absen(a, 8) = ket
                
                a = a + 1
                rs.MoveNext
                Loop
                
                
                grd_absen.Columns(4).FooterText = "Total Data : " & a - 1
                grd_absen.Columns(5).FooterText = "Hadir : " & hadir
                grd_absen.Columns(6).FooterText = "Alpha : " & alpha
                grd_absen.Columns(7).FooterText = "Izin  : " & izin
                grd_absen.Columns(8).FooterText = "Sakit : " & sakit
                
                
                grd_absen.ReBind
                grd_absen.Refresh
End Sub

Private Sub grd_absen_AfterColUpdate(ByVal ColIndex As Integer)

On Error GoTo er_ganti

    If ColIndex = 7 Then
    
        arr_absen(grd_absen.Bookmark, ColIndex) = grd_absen.Columns(ColIndex).Text
            
        Dim sql As String
        Dim rs As New ADODB.Recordset
            
        sql = "update tr_absen set masuk_gak=" & rubah_apsen(arr_absen(grd_absen.Bookmark, ColIndex)) & " where id=" & id_absen
        rs.Open sql, cn
        Exit Sub
    End If
    
    If ColIndex = 8 Then
        
        arr_absen(grd_absen.Bookmark, ColIndex) = grd_absen.Columns(ColIndex).Text
        
        Dim sql1 As String
        Dim rs1 As New ADODB.Recordset
        
        sql1 = "update tr_absen set ket='" & arr_absen(grd_absen.Bookmark, ColIndex) & "' where id=" & id_absen
        rs1.Open sql1, cn
        Exit Sub
        
    End If
            
er_ganti:
        
        Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
            
End Sub

Private Sub grd_absen_Click()
On Error Resume Next
    If arr_absen.UpperBound(1) > 0 Then
        id_absen = arr_absen(grd_absen.Bookmark, 0)
    End If
End Sub

Private Sub grd_absen_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    grd_absen_Click
End Sub
Private Sub msk_jam_GotFocus()
    msk_jam.SelStart = 0
    msk_jam.SelLength = Len(msk_jam)
End Sub

Private Sub msk_tgl_GotFocus()
    msk_tgl.SelStart = 0
    msk_tgl.SelLength = Len(msk_tgl)
End Sub
Private Sub msk_tgl1_GotFocus()
    msk_tgl1.SelStart = 0
    msk_tgl1.SelLength = Len(msk_tgl1)
End Sub

Private Sub txt_nama_GotFocus()
    txt_nama.SelStart = 0
    txt_nama.SelLength = Len(txt_nama)
End Sub
