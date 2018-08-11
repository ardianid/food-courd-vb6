VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "flash.ocx"
Begin VB.MDIForm utama 
   BackColor       =   &H8000000C&
   Caption         =   "Food Courd Central Plaza ==="
   ClientHeight    =   8625
   ClientLeft      =   165
   ClientTop       =   825
   ClientWidth     =   15240
   Icon            =   "utama.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pic_atas 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   0
      ScaleHeight     =   1785
      ScaleWidth      =   15210
      TabIndex        =   0
      Top             =   0
      Width           =   15240
      Begin VB.Timer Timer1 
         Interval        =   600
         Left            =   5880
         Top             =   240
      End
      Begin ShockwaveFlashObjectsCtl.ShockwaveFlash swf 
         Height          =   735
         Left            =   840
         TabIndex        =   1
         Top             =   480
         Width           =   3135
         _cx             =   5530
         _cy             =   1296
         FlashVars       =   ""
         Movie           =   ""
         Src             =   ""
         WMode           =   "Window"
         Play            =   -1  'True
         Loop            =   -1  'True
         Quality         =   "High"
         SAlign          =   ""
         Menu            =   -1  'True
         Base            =   ""
         AllowScriptAccess=   "always"
         Scale           =   "NoBorder"
         DeviceFont      =   0   'False
         EmbedMovie      =   0   'False
         BGColor         =   ""
         SWRemote        =   ""
      End
      Begin VB.Label lbl_user 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User"
         Height          =   195
         Left            =   7440
         TabIndex        =   4
         Top             =   720
         Width           =   330
      End
      Begin VB.Label lbl_jam 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lbl_jam"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   270
         Left            =   8280
         TabIndex        =   3
         Top             =   360
         Width           =   945
      End
      Begin VB.Label lbl_tgl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lbl_tgl"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   270
         Left            =   6240
         TabIndex        =   2
         Top             =   960
         Width           =   945
      End
      Begin VB.Image img_atas 
         Height          =   615
         Left            =   4560
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Menu kry 
      Caption         =   "Master"
      Begin VB.Menu inp_kry 
         Caption         =   "Input Karyawan"
      End
      Begin VB.Menu ar_kry 
         Caption         =   "Data Karyawan"
      End
      Begin VB.Menu data_member 
         Caption         =   "Data Member"
      End
   End
   Begin VB.Menu ctr 
      Caption         =   "&Counter && Barang"
      Begin VB.Menu ip_ctr 
         Caption         =   "Input Data &Counter"
      End
      Begin VB.Menu dt_ctr 
         Caption         =   "&Master Data Counter"
      End
      Begin VB.Menu brg_coun 
         Caption         =   "Input  Data Barang"
      End
   End
   Begin VB.Menu sto 
      Caption         =   " Stock"
      Begin VB.Menu inp_sto 
         Caption         =   "Input Stock"
      End
      Begin VB.Menu br_st 
         Caption         =   "Browse Stock"
      End
      Begin VB.Menu br_stok 
         Caption         =   "Penyesuaian Stok"
      End
      Begin VB.Menu histock 
         Caption         =   "Historical Stock"
      End
   End
   Begin VB.Menu inv 
      Caption         =   "Inventori"
      Begin VB.Menu inp_inv 
         Caption         =   "Input Inventori"
      End
      Begin VB.Menu psstrok 
         Caption         =   "Penyesuaian Stock Inventori"
      End
   End
   Begin VB.Menu pj 
      Caption         =   "Penjualan"
      Begin VB.Menu tr_jual 
         Caption         =   "Transaksi Penjualan"
      End
      Begin VB.Menu bt_trs 
         Caption         =   "&Pembatalan Penjualan"
      End
      Begin VB.Menu br_pj 
         Caption         =   "Browse Penjualan"
      End
      Begin VB.Menu jml_jual 
         Caption         =   "Jumlah Penjualan"
      End
      Begin VB.Menu lp 
         Caption         =   "Laporan Penjualan"
         Begin VB.Menu lp_ksr 
            Caption         =   "Penjualan Perkasir"
         End
         Begin VB.Menu lp_ksr_per 
            Caption         =   "Penjualan Kasir perperiode"
         End
         Begin VB.Menu lap_counter 
            Caption         =   "Laporan PerCounter"
         End
         Begin VB.Menu lap_counter_disc 
            Caption         =   "Laporan PerCounter Berdasarkan Disc"
         End
         Begin VB.Menu lp_tot_jual1 
            Caption         =   "Total Penjualan"
            Begin VB.Menu lp_tot_jual 
               Caption         =   "Total Penjualan"
            End
            Begin VB.Menu tot_jual_per 
               Caption         =   "Total Penjualan Berdasarkan Disc"
            End
         End
      End
      Begin VB.Menu lap_persentase 
         Caption         =   "Persentase"
      End
      Begin VB.Menu slp_bayar 
         Caption         =   "Slip Pembayaran"
      End
   End
   Begin VB.Menu pwd 
      Caption         =   "Password"
      Begin VB.Menu inp_pwd 
         Caption         =   "Input Password"
      End
      Begin VB.Menu brw_pwd 
         Caption         =   "Browse Password"
      End
      Begin VB.Menu hak_aks 
         Caption         =   "Input Hak Akses"
      End
      Begin VB.Menu gt_pwd 
         Caption         =   "Ganti Password"
      End
   End
   Begin VB.Menu j_ker 
      Caption         =   "Jam Kerja"
      Visible         =   0   'False
      Begin VB.Menu dj 
         Caption         =   "Data Jam Kerja"
      End
      Begin VB.Menu pem_tgs_krj 
         Caption         =   "Pembagian Tugas Kerja"
      End
   End
   Begin VB.Menu abs 
      Caption         =   "Absen"
      Visible         =   0   'False
      Begin VB.Menu inp_abs 
         Caption         =   "Input Absen"
      End
      Begin VB.Menu br_ab 
         Caption         =   "Browse Absen"
      End
   End
   Begin VB.Menu by 
      Caption         =   "Biaya Biaya"
      Begin VB.Menu inpt_biaya 
         Caption         =   "Input Biaya"
      End
      Begin VB.Menu br_by 
         Caption         =   "Browse Biaya"
      End
   End
   Begin VB.Menu gj 
      Caption         =   "Penggajian"
      Visible         =   0   'False
      Begin VB.Menu ip_gj 
         Caption         =   "Input Gaji"
      End
      Begin VB.Menu br_gj 
         Caption         =   "Browse Gaji"
      End
      Begin VB.Menu ctk_slip 
         Caption         =   "Cetak Slip Penggajian"
      End
   End
   Begin VB.Menu dtb 
      Caption         =   "Utility"
      Begin VB.Menu set_printer 
         Caption         =   "Seting Printer"
      End
      Begin VB.Menu backp 
         Caption         =   "Backup Database"
      End
   End
   Begin VB.Menu log_off 
      Caption         =   "Log Off"
   End
End
Attribute VB_Name = "utama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim id_wwn As String

Private Sub ar_kry_Click()
If Not (frm Is Nothing) Then
    Unload frm
    Set frm = Nothing
    Set frm = frm_browse_pegawai
    frm.Show
Else
    Set frm = frm_browse_pegawai
    frm.Show
End If
End Sub

Private Sub backp_Click()
If Not (frm Is Nothing) Then
    Unload frm
    Set frm = Nothing
    Set frm = frm_backup
    frm.Show
Else
    Set frm = frm_backup
    frm.Show
End If
End Sub

Private Sub br_ab_Click()
If Not (frm Is Nothing) Then
    Unload frm
    Set frm = Nothing
    Set frm = frm_browse_absen
    frm.Show
Else
    Set frm = frm_browse_absen
    frm.Show
End If
    
End Sub

Private Sub br_bl_Click()
If Not (frm Is Nothing) Then
    Unload frm
    Set frm = Nothing
    Set frm = frm_browse_blg
    frm.Show
Else
    Set frm = frm_browse_blg
    frm.Show
End If
End Sub

Private Sub br_btl_Click()
If Not (frm Is Nothing) Then
    Unload frm
    Set frm = Nothing
    Set frm = frm_browse_btl
    frm.Show
Else
    Set frm = frm_browse_btl
    frm.Show
End If
    
End Sub

Private Sub br_by_Click()
If Not (frm Is Nothing) Then
    Unload frm
    Set frm = Nothing
    Set frm = frm_browse_biaya
    frm.Show
Else
    Set frm = frm_browse_biaya
    frm.Show
End If
End Sub

Private Sub br_gj_Click()
If Not (frm Is Nothing) Then
    Unload frm
    Set frm = Nothing
    Set frm = frm_browse_gaji
    frm.Show
Else
    Set frm = frm_browse_gaji
    frm.Show
End If
End Sub

Private Sub br_pj_Click()
If Not (frm Is Nothing) Then
    Unload frm
    Set frm = Nothing
    Set frm = frm_browse_pnj
    frm.Show
Else
    Set frm = frm_browse_pnj
    frm.Show
End If
    
End Sub

'Private Sub br_slip_Click()
'If Not (frm Is Nothing) Then
'    Unload frm
'    Set frm = Nothing
'    Set frm = frm_browse_slip
'    frm.Show
'Else
'    Set frm = frm_browse_slip
'    frm.Show
'End If
'End Sub

Private Sub br_st_Click()
If Not (frm Is Nothing) Then
    Unload frm
    Set frm = Nothing
    Set frm = frm_browse_tstock
    frm.Show
Else
    Set frm = frm_browse_tstock
    frm.Show
End If
End Sub

Private Sub br_stok_Click()
If Not (frm Is Nothing) Then
    Unload frm
    Set frm = Nothing
    Set frm = frm_pestok
    frm.Show
Else
    Set frm = frm_pestok
    frm.Show
End If
    
End Sub

Private Sub brg_coun_Click()
If Not (frm Is Nothing) Then
    Unload frm
    Set frm = Nothing
    Set frm = frm_bc
    frm.Show
Else
    Set frm = frm_bc
    frm.Show
End If
    
End Sub

Private Sub brw_pwd_Click()
If Not (frm Is Nothing) Then
    Unload frm
    Set frm = Nothing
    Set frm = frm_browse_pwd
    frm.Show
Else
    Set frm = frm_browse_pwd
    frm.Show
End If
    
End Sub

Private Sub ctk_brg_Click()
If Not (frm Is Nothing) Then
    Unload frm
    Set frm = Nothing
    Set frm = frm_cetak_barang
    frm.Show
Else
    Set frm = frm_cetak_barang
    frm.Show
End If
    
End Sub

Private Sub bt_trs_Click()

If Not (frm Is Nothing) Then
    Unload frm
    Set frm = Nothing
    Set frm = frm_btl_tran
    frm.Show
Else
    Set frm = frm_btl_tran
    frm.Show
End If
End Sub

Private Sub ctk_slip_Click()
If Not (frm Is Nothing) Then
    Unload frm
    Set frm = Nothing
    Set frm = frm_slip_penggajian
    frm.Show
Else
    Set frm = frm_slip_penggajian
    frm.Show
End If
End Sub

Private Sub data_member_Click()

Me.Enabled = False

If Not (frm Is Nothing) Then
    Unload frm
    Set frm = Nothing
    Set frm = Frm_Member
    frm.Show
Else
    Set frm = Frm_Member
    frm.Show
End If

End Sub

Private Sub dj_Click()
If Not (frm Is Nothing) Then
    Unload frm
    Set frm = Nothing
    Set frm = frm_jam_kerja
    frm.Show
Else
    Set frm = frm_jam_kerja
    frm.Show
End If
    
End Sub

Private Sub dt_ctr_Click()
If Not (frm Is Nothing) Then
    Unload frm
    Set frm = Nothing
    Set frm = frm_browse_counter
    frm.Show
Else
    Set frm = frm_browse_counter
    frm.Show
End If
    
End Sub

Private Sub gt_pwd_Click()
If Not (frm Is Nothing) Then
    Unload frm
    Set frm = Nothing
    Set frm = frm_ganti_pwd
    frm.Show
Else
    Set frm = frm_ganti_pwd
    frm.Show
End If
    
End Sub

Private Sub hak_aks_Click()
If Not (frm Is Nothing) Then
    Unload frm
    Set frm = Nothing
    Set frm = frm_hak_akses
    frm.Show
Else
    Set frm = frm_hak_akses
    frm.Show
End If
End Sub

Private Sub histock_Click()
If Not (frm Is Nothing) Then
    Unload frm
    Set frm = Nothing
    Set frm = frm_histock
    frm.Show
Else
    Set frm = frm_histock
    frm.Show
End If
    
End Sub

Private Sub inp_abs_Click()
If Not (frm Is Nothing) Then
    Unload frm
    Set frm = Nothing
    Set frm = frm_absen
    frm.Show
Else
    Set frm = frm_absen
    frm.Show
End If
    
End Sub

Private Sub inp_blg_Click()
If Not (frm Is Nothing) Then
    Unload frm
    Set frm = Nothing
    Set frm = frm_biling
    frm.Show
Else
    Set frm = frm_biling
    frm.Show
End If
    
End Sub

Private Sub inp_inv_Click()
If Not (frm Is Nothing) Then
    Unload frm
    Set frm = Nothing
    Set frm = frm_input_inventori
    frm.Show
Else
    Set frm = frm_input_inventori
    frm.Show
End If
End Sub

Private Sub inp_kry_Click()
mdl_karyawan = True
If Not (frm Is Nothing) Then
    Unload frm
    Set frm = Nothing
    Set frm = frm_karyawan_lain
    frm.Show
Else
    Set frm = frm_karyawan_lain
    frm.Show
End If
    
    
End Sub

Private Sub inp_pwd_Click()
If Not (frm Is Nothing) Then
    Unload frm
    Set frm = Nothing
    Set frm = frm_pwd
    frm.Show
Else
    Set frm = frm_pwd
    frm.Show
End If
    
End Sub

Private Sub inp_sto_Click()
mdl_stock = True
If Not (frm Is Nothing) Then
    Unload frm
    Set frm = Nothing
    Set frm = frm_tstock
    frm.Show
Else
    Set frm = frm_tstock
    frm.Show
End If
    
    
End Sub

Private Sub inpt_biaya_Click()
If Not (frm Is Nothing) Then
    Unload frm
    Set frm = Nothing
    Set frm = frm_biaya
    frm.Show
Else
    Set frm = frm_biaya
    frm.Show
End If
End Sub

Private Sub ip_ctr_Click()
    mdl_counter = True
If Not (frm Is Nothing) Then
    Unload frm
    Set frm = Nothing
    Set frm = frm_counter
    frm.Show
Else
    Set frm = frm_counter
    frm.Show
End If
    
End Sub

Private Sub ip_gj_Click()
 If Not (frm Is Nothing) Then
    Unload frm
    Set frm = Nothing
    Set frm = frm_penggajian
    frm.Show
Else
    Set frm = frm_penggajian
    frm.Show
End If
End Sub

Private Sub jml_jual_Click()
If Not (frm Is Nothing) Then
    Unload frm
    Set frm = Nothing
    Set frm = frm_persentase_jual
    frm.Show
Else
    Set frm = frm_persentase_jual
    frm.Show
End If
End Sub

Private Sub lap_bul_Click()
If Not (frm Is Nothing) Then
    Unload frm
    Set frm = Nothing
    Set frm = frm_lap_per_bulan
    frm.Show
Else
    Set frm = frm_lap_per_bulan
    frm.Show
End If
End Sub

Private Sub lap_jual_persentase_Click()
If Not (frm Is Nothing) Then
    Unload frm
    Set frm = Nothing
    Set frm = frm_lap_persentase
    frm.Show
Else
    Set frm = frm_lap_persentase
    frm.Show
End If
End Sub

Private Sub lap_counter_Click()
If Not (frm Is Nothing) Then
    Unload frm
    Set frm = Nothing
    Set frm = frm_lap_jual_perhari
    frm.Show
Else
    Set frm = frm_lap_jual_perhari
    frm.Show
End If
End Sub

Private Sub lap_counter_disc_Click()

If Not (frm Is Nothing) Then
    Unload frm
    Set frm = Nothing
    Set frm = frm_lap_jual_perhari1
    frm.Show
Else
    Set frm = frm_lap_jual_perhari1
    frm.Show
End If

End Sub

Private Sub lap_persentase_Click()

If Not (frm Is Nothing) Then
    Unload frm
    Set frm = Nothing
    Set frm = frm_lap_persentase
    frm.Show
Else
    Set frm = frm_lap_persentase
    frm.Show
End If

End Sub

Private Sub log_off_Click()
    cn.Close
    Set cn = Nothing
    frm_masuk.Show
    Unload Me
End Sub

Private Sub lp_ksr_Click()
If Not (frm Is Nothing) Then
    Unload frm
    Set frm = Nothing
    Set frm = Frm_Sel_Penj_Perkasir
    frm.Show
Else
    Set frm = Frm_Sel_Penj_Perkasir
    frm.Show
End If
End Sub

Private Sub lp_ksr_per_Click()
If Not (frm Is Nothing) Then
    Unload frm
    Set frm = Nothing
    Set frm = Frm_Sel_LapFaktur ' frm_lap_perkasir
    frm.Show
Else
    Set frm = Frm_Sel_LapFaktur 'frm_lap_perkasir
    frm.Show
End If
End Sub

Private Sub lp_pj_h_Click()
If Not (frm Is Nothing) Then
    Unload frm
    Set frm = Nothing
    Set frm = frm_lap_jual_perhari
    frm.Show
Else
    Set frm = frm_lap_jual_perhari
    frm.Show
End If
End Sub

Private Sub lp_tot_jual_Click()
If Not (frm Is Nothing) Then
    Unload frm
    Set frm = Nothing
    Set frm = frm_tot_jual
    frm.Show
Else
    Set frm = frm_tot_jual
    frm.Show
End If
End Sub

Private Sub MDIForm_Load()

Me.Picture = LoadPicture(App.path & "\background utama.jpg")

lbl_tgl.Caption = Format(Date, "Long Date")

cari_wewenang

End Sub

Private Sub cari_wewenang()

On Error GoTo cari

    Dim sql, sql1 As String
    Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
    
    Call nonaktifkan
    
    sql = "select id_wewenang,nama_form from qr_wewenang where id_user=" & id_user
    rs.Open sql, cn, adOpenKeyset
        If Not rs.EOF Then
            
            rs.MoveLast
            rs.MoveFirst
                
                
            Do While Not rs.EOF
                id_wwn = rs!id_wewenang
                
                Call isi(rs!nama_form)
            rs.MoveNext
            Loop
         End If
    rs.Close
    Exit Sub
    
cari:
        Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
    
End Sub

Private Sub nonaktifkan()
    inp_kry.Enabled = False
    ar_kry.Enabled = False
    ip_ctr.Enabled = False
    dt_ctr.Enabled = False
    brg_coun.Enabled = False
   ' ctk_brg.Enabled = False
    inp_sto.Enabled = False
    br_st.Enabled = False
    br_stok.Enabled = False
    histock.Enabled = False
    inp_inv.Enabled = False
    psstrok.Enabled = False
    tr_jual.Enabled = False
    br_pj.Enabled = False
    'lp_pj_h.Enabled = False
    lp_tot_jual.Enabled = False
    inp_pwd.Enabled = False
    brw_pwd.Enabled = False
    hak_aks.Enabled = False
    gt_pwd.Enabled = False
    dj.Enabled = False
    pem_tgs_krj.Enabled = False
    inp_abs.Enabled = False
    br_ab.Enabled = False
'    pbtl_d.Enabled = False
'    br_btl.Enabled = False
'    inp_blg.Enabled = False
'    br_bl.Enabled = False
'    tr_blg.Enabled = False
'    tr_br_bl.Enabled = False
    inpt_biaya.Enabled = False
    br_by.Enabled = False
    ip_gj.Enabled = False
    br_gj.Enabled = False
'    lap_jual_persentase.Enabled = False
'    lap_bul.Enabled = False
    ctk_slip.Enabled = False
    lp_ksr.Enabled = False
    lp_ksr_per.Enabled = False
   ' pj_per.Enabled = False
    tot_jual_per.Enabled = False
    jml_jual.Enabled = False
    slp_bayar.Enabled = False
  '  br_slip.Enabled = False
    backp.Enabled = False
    data_member.Enabled = False
    lap_counter.Enabled = False
    lap_counter_disc.Enabled = False
    lap_persentase.Enabled = False
    
End Sub


Private Sub isi(param)

On Error GoTo er_isi

Dim sql As String
Dim rs As New ADODB.Recordset
    
    If rs.State = adStateOpen Then
        rs.Close
    End If
    
    Select Case param
        Case "Form Karyawan"
            sql = "select * from qr_hak where id_wewenang=" & id_wwn
            rs.Open sql, cn
                
                    If Not rs.EOF Then
                        If rs!tambah = 1 Then
                            inp_kry.Enabled = True
                        End If
                        
                        If rs!edit = 1 Or rs!hapus = 1 Or rs!lap = 1 Then
                            ar_kry.Enabled = True
                        End If
                    End If
                 rs.Close
            
            
        Case "Form Data Counter"
            
            sql = "select * from qr_hak where id_wewenang=" & id_wwn
            rs.Open sql, cn
                
                    If Not rs.EOF Then
                        If rs!tambah = 1 Then
                            ip_ctr.Enabled = True
                        End If
                        
                        If rs!edit = 1 Or rs!hapus = 1 Or rs!lap = 1 Then
                            dt_ctr.Enabled = True
                        End If
                    End If
                 rs.Close
                 
        Case "Form Data Member"

            sql = "select * from qr_hak where id_wewenang=" & id_wwn
            rs.Open sql, cn

                    If Not rs.EOF Then
                        If rs!tambah = 1 Then
                            ip_ctr.Enabled = True
                        End If

                        If rs!edit = 1 Or rs!hapus = 1 Or rs!lap = 1 Then
                            dt_ctr.Enabled = True
                        End If
                        data_member.Enabled = True
                    End If
                 rs.Close
            
            
        Case "Form Data Barang Counter"
            
            sql = "select * from qr_hak where id_wewenang=" & id_wwn
            rs.Open sql, cn
                
                    If Not rs.EOF Then
                        If rs!tambah = 1 Or rs!edit = 1 Or rs!hapus = 1 Then
                            brg_coun.Enabled = True
                        End If
                        
                        If rs!lap = 1 Then
                            'ctk_brg.Enabled = True
                        End If
                    End If
                 rs.Close
            
            
       Case "Form Stock"
            
            sql = "select * from qr_hak where id_wewenang=" & id_wwn
            rs.Open sql, cn
                
                    If Not rs.EOF Then
                        If rs!tambah = 1 Then
                            inp_sto.Enabled = True
                            
                        If rs!lap = 1 Then
                            br_st.Enabled = True
                        End If
                          
                        End If
                    End If
                 rs.Close
            
            
       Case "Form Penyesuaian Stock"
            
            sql = "select * from qr_hak where id_wewenang=" & id_wwn
            rs.Open sql, cn
                
                    If Not rs.EOF Then
                        br_stok.Enabled = True
                    End If
                 rs.Close
            
            
       Case "Form Historical Stock"
            
            sql = "select * from qr_hak where id_wewenang=" & id_wwn
            rs.Open sql, cn
                
                    If Not rs.EOF Then
                        histock.Enabled = True
                    End If
                 rs.Close
            
            
       Case "Form Inventory"
       
            sql = "select * from qr_hak where id_wewenang=" & id_wwn
            rs.Open sql, cn
                
                    If Not rs.EOF Then
                        inp_inv.Enabled = True
                    End If
                 rs.Close
            
            
       Case "Form Penyesuaian Inventory"
            
            sql = "select * from qr_hak where id_wewenang=" & id_wwn
            rs.Open sql, cn
                
                    If Not rs.EOF Then
                        psstrok.Enabled = True
                    End If
                 rs.Close
            
            
       Case "Form Transakai Penjualan"
            
            sql = "select * from qr_hak where id_wewenang=" & id_wwn
            rs.Open sql, cn
                
                    If Not rs.EOF Then
                        tr_jual.Enabled = True
                    End If
                 rs.Close
            
                 
       Case "Form Laporan Penjualan"
            
            sql = "select * from qr_hak where id_wewenang=" & id_wwn
            rs.Open sql, cn
                
                    If Not rs.EOF Then
                        
                      '  lp_pj_h.Enabled = True
                        lp_tot_jual.Enabled = True
'                        lap_jual_persentase.Enabled = True
'                        lap_bul.Enabled = True
                        lp_ksr_per.Enabled = True
                        'pj_per.Enabled = True
                        tot_jual_per.Enabled = True
                        jml_jual.Enabled = True
'                        slp_bayar.Enabled = True
'                        br_slip.Enabled = True
                    End If
                 rs.Close
            
            
       Case "Form Password"
                
            sql = "select * from qr_hak where id_wewenang=" & id_wwn
            rs.Open sql, cn
                
                    If Not rs.EOF Then
                     
                        inp_pwd.Enabled = True
                        brw_pwd.Enabled = True
                        hak_aks.Enabled = True
                        
                    End If
                 rs.Close
            
            
       Case "Form Ganti Password"
            
            sql = "select * from qr_hak where id_wewenang=" & id_wwn
            rs.Open sql, cn
                
                    If Not rs.EOF Then
                        gt_pwd.Enabled = True
                    End If
                 rs.Close
            
            
       Case "Form Data Jam Kerja"
            
            sql = "select * from qr_hak where id_wewenang=" & id_wwn
            rs.Open sql, cn
                
                    If Not rs.EOF Then
                        dj.Enabled = True
                    End If
                 rs.Close
            
            
       Case "Form Pembagian Tugas Kerja"
            
            sql = "select * from qr_hak where id_wewenang=" & id_wwn
            rs.Open sql, cn
                
                    If Not rs.EOF Then
                        pem_tgs_krj.Enabled = True
                    End If
                 rs.Close
            
            
       Case "Form Data Absensi Karyawan"
            
            sql = "select * from qr_hak where id_wewenang=" & id_wwn
            rs.Open sql, cn
                
                    If Not rs.EOF Then
                        
                        If rs!tambah = 1 Then
                            inp_abs.Enabled = True
                        End If
                        
                        If rs!edit = 1 Or rs!hapus = 1 Or rs!lap = 1 Then
                            br_ab.Enabled = True
                        End If
                        
                    End If
                 rs.Close
            
            
       Case "Form Pembatalan Transaksi"
            
            sql = "select * from qr_hak where id_wewenang=" & id_wwn
            rs.Open sql, cn
                
                    If Not rs.EOF Then
                     If rs!tambah = 1 Then
'                        pbtl_d.Enabled = True
                     End If
                     If rs!lap = 1 Then
'                        br_btl.Enabled = True
                     End If
                    End If
                 rs.Close
            
                  
       Case "Form Data Billing"
            
            sql = "select * from qr_hak where id_wewenang=" & id_wwn
            rs.Open sql, cn
                
                    If Not rs.EOF Then
                        
                        If rs!tambah = 1 Then
'                            inp_blg.Enabled = True
                        End If
                        
                        If rs!edit = 1 Or rs!hapus = 1 Or rs!lap = 1 Then
'                            br_bl.Enabled = True
                        End If
                        
                    End If
                 rs.Close
            
            
       Case "Form Transaksi Billing"
            
            sql = "select * from qr_hak where id_wewenang=" & id_wwn
            rs.Open sql, cn
                
                    If Not rs.EOF Then
                        
                        If rs!tambah = 1 Then
'                            tr_blg.Enabled = True
                        End If
                        
                        If rs!hapus = 1 Or rs!lap = 1 Then
'                            tr_br_bl.Enabled = True
                        End If
                        
                    End If
                 rs.Close
            
            
       Case "Form Data Biaya-Biaya"
            
            sql = "select * from qr_hak where id_wewenang=" & id_wwn
            rs.Open sql, cn
                
                    If Not rs.EOF Then
                      
                        inpt_biaya.Enabled = True
                        br_by.Enabled = True
                    End If
                 rs.Close
            
            
       Case "Form Penggajian"
            
            sql = "select * from qr_hak where id_wewenang=" & id_wwn
            rs.Open sql, cn
                
                    If Not rs.EOF Then
                        ip_gj.Enabled = True
                        br_gj.Enabled = True
                        ctk_slip.Enabled = True
                    End If
                 rs.Close
            
            
       Case "Form Browse Penjualan"
            
            sql = "select * from qr_hak where id_wewenang=" & id_wwn
            rs.Open sql, cn
                
                    If Not rs.EOF Then
                        br_pj.Enabled = True
                    End If
                 rs.Close
        Case "Form Lap Perkasir"
            
            sql = "select * from qr_hak where id_wewenang=" & id_wwn
            rs.Open sql, cn
                
                    If Not rs.EOF Then
                        lp_ksr.Enabled = True
                    End If
                 rs.Close
        
        Case "Form Laporan PerCounter"
            
            sql = "select * from qr_hak where id_wewenang=" & id_wwn
            rs.Open sql, cn
                
                    If Not rs.EOF Then
                        lap_counter.Enabled = True
                    End If
                 rs.Close
        
        Case "Form Laporan PerCounter Berdasarkan Disc"
            
            sql = "select * from qr_hak where id_wewenang=" & id_wwn
            rs.Open sql, cn
                
                    If Not rs.EOF Then
                        lap_counter_disc.Enabled = True
                    End If
                 rs.Close
        
        Case "Form Slip Pembayaran"
            
            sql = "select * from qr_hak where id_wewenang=" & id_wwn
            rs.Open sql, cn
                
                    If Not rs.EOF Then
                        slp_bayar.Enabled = True
                    End If
                 rs.Close
        
        Case "Form Persentase"
            
            sql = "select * from qr_hak where id_wewenang=" & id_wwn
            rs.Open sql, cn
                
                    If Not rs.EOF Then
                        lap_persentase.Enabled = True
                    End If
                 rs.Close
        
        Case "Form Backup Database"
            
            sql = "select * from qr_hak where id_wewenang=" & id_wwn
            rs.Open sql, cn
                
                    If Not rs.EOF Then
                        backp.Enabled = True
                    End If
                 rs.Close
        
    End Select
            
Exit Sub

er_isi:
        Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
            
End Sub


Private Sub MDIForm_Resize()

    

    pic_atas.Top = utama.Top
    swf.Top = pic_atas.Top
    swf.Left = pic_atas.Left
    swf.Width = pic_atas.Width
    swf.Height = pic_atas.Height
    swf.Movie = App.path & "\Banner.swf"
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    If Not (frm Is Nothing) Then
        Unload frm
        Set frm = Nothing
    End If
End Sub

Private Sub pbtl_d_Click()
If Not (frm Is Nothing) Then
    Unload frm
    Set frm = Nothing
    Set frm = frm_btl_tran
    frm.Show
Else
    Set frm = frm_btl_tran
    frm.Show
End If
    
End Sub

Private Sub pem_tgs_krj_Click()
If Not (frm Is Nothing) Then
    Unload frm
    Set frm = Nothing
    Set frm = frm_baker
    frm.Show
Else
    Set frm = frm_baker
    frm.Show
End If
    
End Sub

Private Sub pj_per_Click()
If Not (frm Is Nothing) Then
    Unload frm
    Set frm = Nothing
    Set frm = frm_lap_jual_perhari1
    frm.Show
Else
    Set frm = frm_lap_jual_perhari1
    frm.Show
End If
End Sub

Private Sub psstrok_Click()
If Not (frm Is Nothing) Then
    Unload frm
    Set frm = Nothing
    Set frm = frm_penyesuaian_stock
    frm.Show
Else
    Set frm = frm_penyesuaian_stock
    frm.Show
End If
End Sub

Private Sub set_printer_Click()

Me.Enabled = False

'If Not (frm Is Nothing) Then
'    Unload frm
'    Set frm = Nothing
'    Set frm = Frm_SetPrinter
'    frm.Show
'Else
    Set frm = Frm_SetPrinter
    frm.Show
'End If

End Sub


Private Sub slp_bayar_Click()
If Not (frm Is Nothing) Then
    Unload frm
    Set frm = Nothing
    Set frm = frm_slip_pembayaran
    frm.Show
Else
    Set frm = frm_slip_pembayaran
    frm.Show
End If
End Sub

Private Sub Timer1_Timer()
    lbl_jam.Caption = Format(Time, "hh:mm:ss")
End Sub

Private Sub tot_jual_per_Click()
If Not (frm Is Nothing) Then
    Unload frm
    Set frm = Nothing
    Set frm = frm_tot_jual1
    frm.Show
Else
    Set frm = frm_tot_jual1
    frm.Show
End If
End Sub

Private Sub tr_blg_Click()
If Not (frm Is Nothing) Then
    Unload frm
    Set frm = Nothing
    Set frm = frm_tr_biling
    frm.Show
Else
    Set frm = frm_tr_biling
    frm.Show
End If
End Sub

Private Sub tr_br_bl_Click()
If Not (frm Is Nothing) Then
    Unload frm
    Set frm = Nothing
    Set frm = frm_browse_trb
    frm.Show
Else
    Set frm = frm_browse_trb
    frm.Show
End If
End Sub

Private Sub tr_jual_Click()
    
If Not (frm Is Nothing) Then

    Unload frm
    Set frm = Nothing
    Set frm = frm_jual
    frm.Show
    
Else

    Set frm = frm_jual
    frm.Show
    
End If
     
    frm_jual.lbl_user.Caption = utama.lbl_user.Caption
    frm_jual.isi_faktur
     
End Sub
