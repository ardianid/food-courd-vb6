VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_browse_gaji 
   ClientHeight    =   8610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8610
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   8295
      Left            =   120
      ScaleHeight     =   8265
      ScaleWidth      =   14985
      TabIndex        =   0
      Top             =   120
      Width           =   15015
      Begin MSComDlg.CommonDialog cd 
         Left            =   2160
         Top             =   4920
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame Frame3 
         Height          =   855
         Left            =   240
         TabIndex        =   8
         Top             =   7320
         Width           =   14655
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
            Left            =   11520
            TabIndex        =   12
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton cmd_set 
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
            Left            =   9960
            TabIndex        =   11
            Top             =   240
            Width           =   1455
         End
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
            Left            =   13080
            TabIndex        =   10
            Top             =   240
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
            Left            =   240
            TabIndex        =   9
            Top             =   240
            Width           =   1455
         End
      End
      Begin TrueDBGrid60.TDBGrid grd_daftar 
         Height          =   5175
         Left            =   120
         OleObjectBlob   =   "frm_browse_gaji.frx":0000
         TabIndex        =   7
         Top             =   2160
         Width           =   14775
      End
      Begin VB.Frame Frame2 
         Height          =   1455
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   14655
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
            Left            =   13200
            TabIndex        =   17
            Top             =   840
            Width           =   1335
         End
         Begin VB.ComboBox cbo_bulan2 
            Appearance      =   0  'Flat
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
            Left            =   4200
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   840
            Width           =   2415
         End
         Begin VB.ComboBox cbo_bulan1 
            Appearance      =   0  'Flat
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
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   840
            Width           =   2415
         End
         Begin VB.TextBox txt_nama 
            Appearance      =   0  'Flat
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
            Left            =   4200
            TabIndex        =   6
            Top             =   360
            Width           =   2415
         End
         Begin VB.TextBox txt_thn 
            Appearance      =   0  'Flat
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
            Left            =   960
            TabIndex        =   4
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "s/d"
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
            Left            =   3600
            TabIndex        =   15
            Top             =   840
            Width           =   315
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bulan"
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
            TabIndex        =   13
            Top             =   840
            Width           =   585
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Karyawan"
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
            Left            =   2280
            TabIndex        =   5
            Top             =   360
            Width           =   1725
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Thn"
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
            Top             =   360
            Width           =   375
         End
      End
      Begin VB.Line Line1 
         X1              =   480
         X2              =   14640
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pencarian"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Width           =   1755
      End
   End
End
Attribute VB_Name = "frm_browse_gaji"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arr_daftar As New XArrayDB
Dim id_grid As String

Private Sub cmd_cetak_Click()

On Error GoTo er_printer

    With grd_daftar.PrintInfo
        
        .PageHeaderFont.Bold = True
        .PageHeaderFont.Italic = True
        .PageHeaderFont.Size = 10
        .PageHeader = "Laporan Data Penggajian Karyawan"
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

Private Sub cmd_export_Click()
    On Error Resume Next
        cd.ShowSave
        grd_daftar.ExportToFile cd.FileName, False
End Sub

Private Sub cmd_hapus_Click()

On Error GoTo er_hps

Dim sql, sql1 As String
Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
    
    If arr_daftar.UpperBound(1) = 0 Then
        Exit Sub
    End If
    
    If MsgBox("Yakin akan hapus data gaji karyawan  ", vbYesNo + vbQuestion, "Paesan") = vbNo Then
        Exit Sub
    End If
    
    sql = "select id from tbl_gaji where id= " & id_grid
    rs.Open sql, cn
        If Not rs.EOF Then
            
            sql1 = "delete from tbl_gaji where id=" & id_grid
            rs1.Open sql1, cn
            
        Else
            
            MsgBox ("Data yang akan dihapus tidak ditemukan")
            
        End If
    rs.Close
    cmd_tampil_Click
    Exit Sub
    
er_hps:
    Dim ps
        ps = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
End Sub

Private Sub cmd_set_Click()
On Error GoTo er_jangan
    
    With grd_daftar.PrintInfo
        .PageSetup
    End With
    Exit Sub
    
er_jangan:
    Dim o
        o = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
    
End Sub

Private Sub cmd_tampil_Click()
    isi_daftar
End Sub

Private Sub Form_Load()

grd_daftar.Array = arr_daftar

kosong_daftar

txt_thn.Text = Year(Now)

isi_combo1
cbo_bulan1.ListIndex = 0
isi_combo2
cbo_bulan2.ListIndex = 0

End Sub

Private Sub isi_daftar()
    
On Error GoTo er_isi
    
    Dim sql As String
    Dim rsd As New ADODB.Recordset
        
    kosong_daftar
    
    sql = "select * from qr_gaji"
        
        If txt_thn.Text <> "" Or txt_nama.Text <> "" Or cbo_bulan1.Text <> "Semua" Or cbo_bulan2.Text <> "Semua" Then
            
            sql = sql & " where"
            
            If txt_thn.Text <> "" Then
                sql = sql & " thn=" & Trim(txt_thn.Text) & ""
            End If
            
            If txt_nama.Text <> "" And txt_thn.Text = "" Then
                sql = sql & " nama_karyawan like '%" & Trim(txt_nama.Text) & "%'"
            End If
            
            If txt_nama.Text <> "" And txt_thn.Text <> "" Then
                sql = sql & " and nama_karyawan like '%" & Trim(txt_nama.Text) & "%'"
            End If
            
            If cbo_bulan1.Text <> "Semua" And cbo_bulan2.Text <> "Semua" And txt_nama.Text = "" And txt_thn.Text = "" Then
                sql = sql & " bulan >=" & bulan(cbo_bulan1.Text) & " and bulan <= " & bulan(cbo_bulan2.Text) & ""
            End If
            
            If cbo_bulan1.Text <> "Semua" And cbo_bulan2.Text <> "Semua" And (txt_nama.Text <> "" Or txt_thn.Text <> "") Then
                sql = sql & " and bulan >=" & bulan(cbo_bulan1.Text) & " and bulan <= " & bulan(cbo_bulan2.Text) & ""
            End If
       End If
       
       sql = sql & " order by tgl,bulan"
       rsd.Open sql, cn, adOpenKeyset
            If Not rsd.EOF Then
                
                rsd.MoveLast
                rsd.MoveFirst
                
                lanjut rsd
            End If
       rsd.Close
    
    Exit Sub
    
er_isi:
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
    
End Sub

Private Sub lanjut(rsd As Recordset)
    Dim id_dia, tgl, bl, nm, gaji_pokok, tunjangan, lain, potongan, total, user As String
    Dim a As Long
        
        a = 1
            Do While Not rsd.EOF
            arr_daftar.ReDim 1, a, 0, 11
            grd_daftar.ReBind
            grd_daftar.Refresh
                
                id_dia = rsd!id
                
                If Not IsNull(rsd!tgl) Then
                    tgl = rsd!tgl
                Else
                    tgl = ""
                End If
                
                If Not IsNull(rsd!bulan) Then
                    bl = balik_bulan(rsd!bulan)
                Else
                    bl = ""
                End If
                
                If Not IsNull(rsd!nama_karyawan) Then
                    nm = rsd!nama_karyawan
                Else
                    nm = ""
                End If
                
                If Not IsNull(rsd!gaji_pokok) Then
                    gaji_pokok = rsd!gaji_pokok
                Else
                    gaji_pokok = 0
                End If
                
                If Not IsNull(rsd!tunjangan) Then
                    tunjangan = rsd!tunjangan
                Else
                    tunjangan = 0
                End If
                
                If Not IsNull(rsd!lain_lain) Then
                    lain = rsd!lain_lain
                Else
                    lain = 0
                End If
                
                If Not IsNull(rsd!potongan) Then
                    potongan = rsd!potongan
                Else
                    potongan = 0
                End If
                
                If Not IsNull(rsd!gaji_diterima) Then
                    total = rsd!gaji_diterima
                Else
                    total = 0
                End If
                
                If Not IsNull(rsd!nama_user) Then
                    user = rsd!nama_user
                Else
                    user = ""
                End If
         
         arr_daftar(a, 0) = id_dia
         arr_daftar(a, 1) = a
         arr_daftar(a, 2) = tgl
         arr_daftar(a, 3) = bl
         arr_daftar(a, 4) = nm
         arr_daftar(a, 5) = gaji_pokok
         arr_daftar(a, 6) = tunjangan
         arr_daftar(a, 7) = lain
         arr_daftar(a, 8) = potongan
         arr_daftar(a, 9) = total
         arr_daftar(a, 10) = user
         a = a + 1
         rsd.MoveNext
         Loop
         
         grd_daftar.ReBind
         grd_daftar.Refresh
                
End Sub

Private Sub isi_combo1()
    
    With cbo_bulan1
         .AddItem "Semua"
         .AddItem "Januari"
         .AddItem "Februari"
         .AddItem "Maret"
         .AddItem "April"
         .AddItem "Mei"
         .AddItem "Juni"
         .AddItem "Juli"
         .AddItem "Agustus"
         .AddItem "September"
         .AddItem "Oktober"
         .AddItem "Nopember"
         .AddItem "Desember"
    End With

End Sub

Private Sub isi_combo2()

    With cbo_bulan2
         .AddItem "Semua"
         .AddItem "Januari"
         .AddItem "Februari"
         .AddItem "Maret"
         .AddItem "April"
         .AddItem "Mei"
         .AddItem "Juni"
         .AddItem "Juli"
         .AddItem "Agustus"
         .AddItem "September"
         .AddItem "Oktober"
         .AddItem "Nopember"
         .AddItem "Desember"
    End With
    
End Sub

Private Sub kosong_daftar()
    arr_daftar.ReDim 0, 0, 0, 0
    grd_daftar.ReBind
    grd_daftar.Refresh
End Sub

Private Sub grd_daftar_Click()
On Error Resume Next
    If arr_daftar.UpperBound(1) > 0 Then
       id_grid = arr_daftar(grd_daftar.Bookmark, 0)
    End If
End Sub

Private Sub grd_daftar_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    grd_daftar_Click
End Sub

Private Sub txt_nama_GotFocus()
    txt_nama.SelStart = 0
    txt_nama.SelLength = Len(txt_nama)
End Sub

Private Sub txt_thn_GotFocus()
    txt_thn.SelStart = 0
    txt_thn.SelLength = Len(txt_thn)
End Sub

Private Sub txt_thn_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then
    Beep
    KeyAscii = 0
End If
End Sub
