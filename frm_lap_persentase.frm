VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_lap_persentase 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   8625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8625
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cd 
      Left            =   2400
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      ScaleHeight     =   825
      ScaleWidth      =   14985
      TabIndex        =   7
      Top             =   7560
      Width           =   15015
      Begin VB.CommandButton cmd_export 
         Caption         =   "Export"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3720
         TabIndex        =   10
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmd_cetak 
         Caption         =   "Cetak"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1920
         TabIndex        =   9
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmd_setup 
         Caption         =   "Page Setup"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   1695
      End
   End
   Begin TrueDBGrid60.TDBGrid grd_daftar 
      Height          =   6375
      Left            =   120
      OleObjectBlob   =   "frm_lap_persentase.frx":0000
      TabIndex        =   6
      Top             =   1080
      Width           =   15015
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      ScaleHeight     =   825
      ScaleWidth      =   14985
      TabIndex        =   0
      Top             =   120
      Width           =   15015
      Begin VB.CheckBox cek_lain 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Perhitungkan Disc"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6960
         TabIndex        =   11
         Top             =   240
         Width           =   3735
      End
      Begin VB.CommandButton cmd_tampil 
         Caption         =   "Tampil"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   13440
         TabIndex        =   3
         Top             =   120
         Width           =   1455
      End
      Begin MSMask.MaskEdBox msk_tgl1 
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Top             =   240
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_tgl2 
         Height          =   375
         Left            =   3240
         TabIndex        =   2
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Line Line2 
         X1              =   10800
         X2              =   10800
         Y1              =   0
         Y2              =   840
      End
      Begin VB.Line Line1 
         X1              =   6720
         X2              =   6720
         Y1              =   0
         Y2              =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "s/d"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2760
         TabIndex        =   5
         Top             =   240
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Periode"
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
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   870
      End
   End
End
Attribute VB_Name = "frm_lap_persentase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arr_daftar As New XArrayDB

Private Sub cmd_cetak_Click()
 
    On Error GoTo er_printer

    With grd_daftar.PrintInfo
        
        .PageFooterFont.Name = "Arial"
        .PageHeaderFont.Size = 12
        .PageHeader = "TOTAL LAPORAN PENJUALAN & PERSENTASE \t\tPeriode  : " & Trim(msk_tgl1.Text) & " s/d " & Trim(msk_tgl2.Text)
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

Private Sub cmd_setup_Click()
    On Error GoTo er_setup
        With grd_daftar.PrintInfo
            .PageSetup
        End With
        Exit Sub
        
er_setup:
        Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
End Sub

Private Sub Cmd_Tampil_Click()
    isi
End Sub

Private Sub Form_Load()

    grd_daftar.Array = arr_daftar
    
    kosong_daftar
    
End Sub

Private Sub kosong_daftar()
    arr_daftar.ReDim 0, 0, 0, 0
    grd_daftar.ReBind
    grd_daftar.Refresh
End Sub

Private Sub isi()

On Error GoTo er_isi

    Dim sql, sql1, sql2 As String
    Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset, rs2 As New ADODB.Recordset
    Dim nilai As Double, ppn As Double, tot As Double, tot_jual As Double, persen
    Dim a As Long
    Dim jml_jual As Double, jml_nilai As Double, jml_ppn As Double, jml_tot As Double
    Dim kd_counter, nm_counter As String
        
        
        If msk_tgl1.Text = "__/__/____" And msk_tgl2.Text = "__/__/____" Then
            MsgBox ("Periode laporan hrs diisi semua")
            msk_tgl1.SetFocus
            Exit Sub
        End If
        
        kosong_daftar
        
        sql = "select distinct(kode_counter)as kd_c from qr_penjualan_sebenarnya where"
        sql = sql & " tgl >= datevalue('" & Trim(msk_tgl1.Text) & "') and tgl <= datevalue('" & Trim(msk_tgl2.Text) & "')"
        sql = sql & " order by kode_counter"
        
        rs.Open sql, cn, adOpenKeyset
            If Not rs.EOF Then
                
                a = 1
                jml_jual = 0
                jml_nilai = 0
                jml_ppn = 0
                jml_tot = 0
                
                Do While Not rs.EOF
                    
                    sql1 = "select nama_counter,presentasi_p from tbl_counter where kode='" & Trim(rs!kd_c) & "'"
                    rs1.Open sql1, cn
                        If Not rs1.EOF Then
                        
                         If cek_lain.Value = vbUnchecked Then
                            sql2 = "select sum(harga_sebenarnya) as benar from qr_penjualan_sebenarnya where"
                            sql2 = sql2 & " kode_counter='" & Trim(rs!kd_c) & "' and tgl >= datevalue('" & Trim(msk_tgl1.Text) & "') and tgl <= datevalue('" & Trim(msk_tgl2.Text) & "')"
                         End If
                                
                         If cek_lain.Value = vbChecked Then
                           sql2 = "select sum(total_harga) as benar from qr_penjualan_sebenarnya where"
                           sql2 = sql2 & " kode_counter='" & Trim(rs!kd_c) & "' and tgl >= datevalue('" & Trim(msk_tgl1.Text) & "') and tgl <= datevalue('" & Trim(msk_tgl2.Text) & "')"
                         End If
                                
                                rs2.Open sql2, cn
                                    If Not rs2.EOF Then
                                        
                                        kd_counter = rs!kd_c
                                        nm_counter = rs1!nama_counter
                                        
                                        tot_jual = rs2!benar
                                        persen = Mid(rs1!presentasi_p, 1, Len(rs1!presentasi_p) - 1)
                                        nilai = (CDbl(tot_jual) * CDbl(persen)) / 100
                                        ppn = (CDbl(tot_jual) - CDbl(nilai)) * (10 / 100)
                                        tot = CDbl(nilai) + CDbl(ppn)
                                        
                                        arr_daftar.ReDim 1, a, 0, 9
                                        grd_daftar.ReBind
                                        grd_daftar.Refresh
                                            
                                            jml_jual = CDbl(jml_jual) + CDbl(tot_jual)
                                            jml_nilai = CDbl(jml_nilai) + CDbl(nilai)
                                            jml_ppn = CDbl(jml_ppn) + CDbl(ppn)
                                            jml_tot = CDbl(jml_tot) + CDbl(tot)
                                            
                                            arr_daftar(a, 0) = a
                                            arr_daftar(a, 1) = kd_counter
                                            arr_daftar(a, 2) = nm_counter
                                            arr_daftar(a, 3) = Format(tot_jual, "###,###,###")
                                            arr_daftar(a, 4) = persen
                                            arr_daftar(a, 5) = Format(nilai, "###,###,###")
                                            arr_daftar(a, 6) = Format(ppn, "###,###,###")
                                            arr_daftar(a, 7) = Format(tot, "###,###,###")
                                            
                                    a = a + 1
                                    End If
                                rs2.Close
                        Else
                            MsgBox ("Nama Counter Tidak ditemukan")
                            Exit Sub
                        End If
                      rs1.Close
                 rs.MoveNext
                 Loop
            End If
         rs.Close
                  grd_daftar.Columns(2).FooterText = "TOTAL"
                  grd_daftar.Columns(3).FooterText = Format(jml_jual, "###,###,###")
                  grd_daftar.Columns(5).FooterText = Format(jml_nilai, "###,###,###")
                  grd_daftar.Columns(6).FooterText = Format(jml_ppn, "###,###,###")
                  grd_daftar.Columns(7).FooterText = Format(jml_tot, "###,###,###")
                  
                  grd_daftar.ReBind
                  grd_daftar.Refresh
                  Exit Sub
Exit Sub

er_isi:
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
        
End Sub

Private Sub msk_tgl1_GotFocus()
    msk_tgl1.SelStart = 0
    msk_tgl1.SelLength = Len(msk_tgl1)
End Sub

Private Sub msk_tgl2_GotFocus()
    msk_tgl2.SelStart = 0
    msk_tgl2.SelLength = Len(msk_tgl2)
End Sub
