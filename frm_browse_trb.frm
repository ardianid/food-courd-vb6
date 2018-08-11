VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_browse_trb 
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8955
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   8295
      Left            =   240
      ScaleHeight     =   8265
      ScaleWidth      =   14865
      TabIndex        =   0
      Top             =   120
      Width           =   14895
      Begin VB.PictureBox Picture3 
         Height          =   735
         Left            =   240
         ScaleHeight     =   675
         ScaleWidth      =   14355
         TabIndex        =   16
         Top             =   7440
         Width           =   14415
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
            TabIndex        =   20
            Top             =   120
            Width           =   1455
         End
         Begin VB.CommandButton cmd_Cetak 
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
            TabIndex        =   19
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
            TabIndex        =   18
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
            Left            =   12960
            TabIndex        =   17
            Top             =   120
            Width           =   1335
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   4695
         Left            =   240
         ScaleHeight     =   4665
         ScaleWidth      =   14385
         TabIndex        =   14
         Top             =   2640
         Width           =   14415
         Begin MSComDlg.CommonDialog cd 
            Left            =   1680
            Top             =   2160
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin TrueDBGrid60.TDBGrid grd_daftar 
            Height          =   4455
            Left            =   120
            OleObjectBlob   =   "frm_browse_trb.frx":0000
            TabIndex        =   15
            Top             =   120
            Width           =   14175
         End
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
         Left            =   13200
         TabIndex        =   13
         Top             =   1920
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
         TabIndex        =   12
         Top             =   1920
         Width           =   4215
      End
      Begin VB.TextBox txt_kode 
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
         TabIndex        =   10
         Top             =   1320
         Width           =   2775
      End
      Begin MSMask.MaskEdBox msk_tgl 
         Height          =   375
         Left            =   2400
         TabIndex        =   6
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Frame Frame1 
         Caption         =   "Data Billing yang akan ditampilkan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   10800
         TabIndex        =   1
         Top             =   0
         Width           =   3975
         Begin VB.OptionButton opt_air 
            Caption         =   "Air"
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
            Left            =   2400
            TabIndex        =   3
            Top             =   480
            Width           =   975
         End
         Begin VB.OptionButton opt_listrik 
            Caption         =   "Listrik"
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
            Left            =   600
            TabIndex        =   2
            Top             =   480
            Width           =   1455
         End
      End
      Begin MSMask.MaskEdBox msk_tgl1 
         Height          =   375
         Left            =   4560
         TabIndex        =   8
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Line Line2 
         X1              =   240
         X2              =   14640
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Nama Counter"
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
         Left            =   600
         TabIndex        =   11
         Top             =   2040
         Width           =   1515
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Counter"
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
         Left            =   600
         TabIndex        =   9
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "S/d"
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
         Left            =   3960
         TabIndex        =   7
         Top             =   720
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl."
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
         Left            =   600
         TabIndex        =   5
         Top             =   720
         Width           =   375
      End
      Begin VB.Line Line1 
         X1              =   480
         X2              =   10800
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
         TabIndex        =   4
         Top             =   120
         Width           =   1140
      End
   End
End
Attribute VB_Name = "frm_browse_trb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arr_daftar As New XArrayDB
Dim sql As String
Dim id_biling As String

Private Sub cmd_cetak_Click()

On Error GoTo er_printer

    With grd_daftar.PrintInfo
        
        .PageHeaderFont.Bold = True
        .PageHeaderFont.Italic = True
        .PageHeaderFont.Size = 10
        
       If opt_listrik.Value = True Then
        .PageHeader = "Laporan Billing Listrik"
       End If
       
       If opt_air.Value = True Then
        .PageHeader = "Laporan Billing Air"
       End If
       
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
Dim sql2, sql1 As String
Dim rs2 As New ADODB.Recordset, rs1 As New ADODB.Recordset
    
On Error GoTo er_h
    
    If arr_daftar.UpperBound(1) = 0 Then
        Exit Sub
    End If
    
    If MsgBox("Yakin Akan dihapus", vbYesNo + vbQuestion, "Pesan") = vbNo Then
        Exit Sub
    End If
    
    If opt_listrik.Value = True Then
        sql2 = "select id from tr_biling_listrik where id=" & id_biling
    ElseIf opt_air.Value = True Then
        sql2 = "select id from tr_biling_air where id=" & id_biling
    End If
        
        rs2.Open sql2, cn
            If Not rs2.EOF Then
                
                If opt_listrik.Value = True Then
                    sql1 = "delete from tr_biling_listrik where id=" & id_biling
                ElseIf opt_air.Value = True Then
                    sql1 = "delete from tr_biling_air where id=" & id_biling
                End If
                    
                    rs1.Open sql1, cn
            Else
                MsgBox ("Data yang akan dihapus tidak ditemukan")
            End If
        rs2.Close
        cmd_tampil_Click
        Exit Sub
        
er_h:
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
End Sub

Private Sub cmd_setup_Click()
On Error GoTo er_set
    With grd_daftar.PrintInfo
        .PageSetup
    End With
    Exit Sub
    
er_set:
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear

End Sub

Private Sub cmd_tampil_Click()
    
    kosong_daftar
    
    Call isi_daftar
End Sub

Private Sub Form_Load()

    grd_daftar.Array = arr_daftar
    
    opt_listrik.Value = True
    
    kosong_daftar
    
    Call cari_wewenang("Form Transaksi Billing")
        
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

Private Sub kosong_daftar()
    arr_daftar.ReDim 0, 0, 0, 0
    grd_daftar.ReBind
    grd_daftar.Refresh
End Sub

Private Sub isi_daftar()

On Error GoTo er_daftar

Dim rs As New ADODB.Recordset
        
    If opt_listrik.Value = True Then
        sql = "select * from qr_biling_listrik"
    ElseIf opt_air.Value = True Then
        sql = "select * from qr_biling_air"
    End If
            
        If msk_tgl.Text <> "__/__/____" Or msk_tgl1.Text <> "__/__/____" Or txt_kode.Text <> "" Or txt_nama.Text <> "" Then
            sql = sql & " where"
            
            If msk_tgl.Text <> "__/__/____" And msk_tgl1.Text = "__/__/____" Then
                sql = sql & " tgl=datevalue('" & Trim(msk_tgl.Text) & "')"
            End If
            
            If msk_tgl1.Text <> "__/__/____" And msk_tgl.Text = "__/__/____" Then
                sql = sql & " tgl=datevalue('" & Trim(msk_tgl1.Text) & "')"
            End If
            
            If msk_tgl1.Text <> "__/__/____" And msk_tgl.Text <> "__/__/____" Then
                sql = sql & " tgl >= datevalue('" & Trim(msk_tgl.Text) & "') and tgl <= datevalue('" & Trim(msk_tgl1.Text) & "')"
            End If
            
            If txt_kode.Text <> "" And msk_tgl1.Text = "__/__/____" And msk_tgl.Text = "__/__/____" Then
                sql = sql & " kode like '%" & Trim(txt_kode.Text) & "%'"
            End If
            
            If txt_kode.Text <> "" And (msk_tgl1.Text <> "__/__/____" Or msk_tgl.Text <> "__/__/____") Then
                sql = sql & " and kode like '%" & Trim(txt_kode.Text) & "%'"
            End If
            
            If txt_nama.Text <> "" And txt_kode.Text = "" And msk_tgl1.Text = "__/__/____" And msk_tgl.Text = "__/__/____" Then
                sql = sql & " nama_counter like '%" & Trim(txt_nama.Text) & "%'"
            End If
            
            If txt_nama.Text <> "" And (txt_kode.Text <> "" Or msk_tgl1.Text <> "__/__/____" Or msk_tgl.Text <> "__/__/____") Then
                sql = sql & " and nama_counter like '%" & Trim(txt_nama.Text) & "%'"
            End If
            
        End If
             
        rs.Open sql, cn, adOpenKeyset
            If Not rs.EOF Then
                
                rs.MoveLast
                rs.MoveFirst
                
                lanjut rs
            End If
        rs.Close
            
        Exit Sub
        
er_daftar:
            Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
            
End Sub

Private Sub lanjut(rs As Recordset)
    Dim id_dia, tgl, kode, nama, pakai, bayar, user As String
    Dim a As Long
        
        a = 1
            Do While Not rs.EOF
                arr_daftar.ReDim 1, a, 0, 9
                grd_daftar.ReBind
                grd_daftar.Refresh
                    
                    If opt_listrik.Value = True Then
                        id_dia = rs("id_biling_listrik")
                    ElseIf opt_air.Value = True Then
                        id_dia = rs("id_biling_air")
                    End If
                    
                    If Not IsNull(rs("tgl")) Then
                        tgl = rs("tgl")
                    Else
                        tgl = ""
                    End If
                    
                    If Not IsNull(rs("kode")) Then
                        kode = rs("kode")
                    Else
                        kode = ""
                    End If
                    
                    If Not IsNull(rs("nama_counter")) Then
                        nama = rs("nama_counter")
                    Else
                        nama = ""
                    End If
                    
                    If Not IsNull(rs("pakai")) Then
                        pakai = rs("pakai")
                    Else
                        pakai = ""
                    End If
                    
                    If Not IsNull(rs("harga")) Then
                        bayar = rs("harga")
                    Else
                        bayar = ""
                    End If
                    
                    If Not IsNull(rs("nama_user")) Then
                        user = rs("nama_user")
                    Else
                        user = ""
                    End If
                    
                arr_daftar(a, 0) = id_dia
                arr_daftar(a, 1) = a
                arr_daftar(a, 2) = tgl
                arr_daftar(a, 3) = kode
                arr_daftar(a, 4) = nama
                arr_daftar(a, 5) = pakai
                arr_daftar(a, 6) = bayar
                arr_daftar(a, 7) = user
                
            a = a + 1
            rs.MoveNext
            Loop
            grd_daftar.ReBind
            grd_daftar.Refresh
End Sub

Private Sub grd_daftar_Click()
On Error Resume Next
    If arr_daftar.UpperBound(1) > 0 Then
        id_biling = arr_daftar(grd_daftar.Bookmark, 0)
    End If
End Sub

Private Sub grd_daftar_HeadClick(ByVal ColIndex As Integer)

On Error GoTo er_head

    Dim sql1 As String
    Dim rs As New ADODB.Recordset
        
        If sql = "" Then
            Exit Sub
        End If
        
        If arr_daftar.UpperBound(1) = 0 Then
            Exit Sub
        End If
        
        sql1 = ""
        sql1 = sql1 & sql
        
        Select Case ColIndex
            Case 2
                sql1 = sql1 & " order by tgl"
            Case 3
                sql1 = sql1 & " order by kode"
            Case 4
                sql1 = sql1 & " order by nama_counter"
            Case 5
                sql1 = sql1 & " order by pakai"
            Case 6
                sql1 = sql1 & " order by harga"
            Case 7
                sql1 = sql1 & " order by nama_user"
        End Select
            
            rs.Open sql1, cn, adOpenKeyset
            If Not rs.EOF Then
                
                rs.MoveLast
                rs.MoveFirst
                
                lanjut rs
            End If
            rs.Close
        
        Exit Sub
        
er_head:
        Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
        
End Sub

Private Sub grd_daftar_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    grd_daftar_Click
End Sub

Private Sub opt_air_Click()
    cmd_tampil_Click
End Sub

Private Sub opt_listrik_Click()
    cmd_tampil_Click
End Sub
