VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_browse_biaya 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   11670
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6135
      Left            =   120
      ScaleHeight     =   6105
      ScaleWidth      =   11385
      TabIndex        =   0
      Top             =   120
      Width           =   11415
      Begin MSComDlg.CommonDialog cd 
         Left            =   3480
         Top             =   3840
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
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
         Left            =   9840
         TabIndex        =   15
         Top             =   1560
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
         Left            =   9840
         TabIndex        =   14
         Top             =   5520
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
         Left            =   8400
         TabIndex        =   13
         Top             =   5520
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
         Left            =   6960
         TabIndex        =   12
         Top             =   5520
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
         Left            =   120
         TabIndex        =   11
         Top             =   5520
         Width           =   1455
      End
      Begin TrueDBGrid60.TDBGrid grd_daftar 
         Height          =   3255
         Left            =   120
         OleObjectBlob   =   "frm_browse_biaya.frx":0000
         TabIndex        =   10
         Top             =   2160
         Width           =   11175
      End
      Begin VB.TextBox txt_ket 
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
         Left            =   1560
         TabIndex        =   9
         Top             =   1560
         Width           =   2895
      End
      Begin VB.TextBox txt_no_bukti 
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
         Left            =   1560
         TabIndex        =   7
         Top             =   1080
         Width           =   2895
      End
      Begin MSMask.MaskEdBox msk_tgl1 
         Height          =   375
         Left            =   1560
         TabIndex        =   3
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox msk_tgl2 
         Height          =   375
         Left            =   3840
         TabIndex        =   5
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
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
      Begin VB.Label Label5 
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
         Left            =   480
         TabIndex        =   8
         Top             =   1560
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Bukti"
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
         Left            =   480
         TabIndex        =   6
         Top             =   1080
         Width           =   885
      End
      Begin VB.Label Label3 
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
         Left            =   3360
         TabIndex        =   4
         Top             =   600
         Width           =   315
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
         Left            =   480
         TabIndex        =   2
         Top             =   600
         Width           =   375
      End
      Begin VB.Line Line1 
         X1              =   360
         X2              =   11160
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pencarian"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Width           =   1350
      End
   End
End
Attribute VB_Name = "frm_browse_biaya"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arr_daftar As New XArrayDB
Dim no_buk As String

Private Sub kosong_daftar()
    arr_daftar.ReDim 0, 0, 0, 0
    grd_daftar.ReBind
    grd_daftar.Refresh
End Sub

Private Sub isi_daftar()

On Error GoTo isi_daftar

    Dim sql As String
    Dim rs As New ADODB.Recordset
        
        kosong_daftar
        
        sql = "select * from qr_biaya"
        
        If msk_tgl1.Text <> "__/__/____" Or msk_tgl2.Text <> "__/__/____" Or txt_no_bukti.Text <> "" Or txt_ket.Text <> "" Then
            
            sql = sql & " where"
            
            If msk_tgl1.Text <> "__/__/____" And msk_tgl2.Text = "__/__/____" Then
                sql = sql & " tgl= datevalue('" & Trim(msk_tgl1.Text) & "')"
            End If
            
            If msk_tgl2.Text <> "__/__/____" And msk_tgl1.Text = "__/__/____" Then
                sql = sql & " tgl= datevalue('" & Trim(msk_tgl2.Text) & "')"
            End If
            
            If msk_tgl2.Text <> "__/__/____" And msk_tgl1.Text <> "__/__/____" Then
                sql = sql & " tgl >= datevalue('" & Trim(msk_tgl1.Text) & "') and tgl <= datevalue('" & Trim(msk_tgl2.Text) & "')"
            End If
            
            If txt_no_bukti.Text <> "" And msk_tgl2.Text = "__/__/____" And msk_tgl1.Text = "__/__/____" Then
                sql = sql & " no_bukti like '%" & Trim(txt_no_bukti.Text) & "%'"
            End If
            
            If txt_no_bukti.Text <> "" And (msk_tgl2.Text <> "__/__/____" Or msk_tgl1.Text <> "__/__/____") Then
                sql = sql & " and no_bukti like '%" & Trim(txt_no_bukti.Text) & "%'"
            End If
            
            If txt_ket.Text <> "" And txt_no_bukti.Text = "" And msk_tgl2.Text = "__/__/____" And msk_tgl1.Text = "__/__/____" Then
                sql = sql & " ket like '%" & Trim(txt_ket.Text) & "%'"
            End If
            
            If txt_ket.Text <> "" And (txt_no_bukti.Text <> "" Or msk_tgl2.Text <> "__/__/____" Or msk_tgl1.Text <> "__/__/____") Then
                sql = sql & " and ket like '%" & Trim(txt_ket.Text) & "%'"
            End If
        End If
        
        sql = sql & " order by tgl,no_bukti"
        rs.Open sql, cn, adOpenKeyset
            If Not rs.EOF Then
                
                rs.MoveLast
                rs.MoveFirst
                    
                    lanjut rs
            End If
        rs.Close
            
       Exit Sub
       
isi_daftar:
       Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
            
End Sub

Private Sub lanjut(rs As Recordset)
    Dim tgl, no_bukti, ket, biaya, user As String
    Dim a, b As Long
    Dim tot_biaya As Double
        a = 1
        b = 1
        tot_biaya = 0
            Do While Not rs.EOF
                arr_daftar.ReDim 1, a, 0, 7
                grd_daftar.ReBind
                grd_daftar.Refresh
                    
                    If Not IsNull(rs("tgl")) Then
                        tgl = rs("tgl")
                    Else
                        tgl = ""
                    End If
                    
                    If Not IsNull(rs("no_bukti")) Then
                        no_bukti = rs("no_bukti")
                    Else
                        no_bukti = ""
                    End If
                    
                    If Not IsNull(rs("ket")) Then
                        ket = rs("ket")
                    Else
                        ket = ""
                    End If
                    
                    If Not IsNull(rs("biaya")) Then
                        biaya = rs("biaya")
                    Else
                        biaya = 0
                    End If
                    
                    If Not IsNull(rs("nama_user")) Then
                        user = rs("nama_user")
                    Else
                        user = ""
                    End If
                    
                    If a > 1 Then
                        If no_bukti <> arr_daftar(a - 1, 2) Then
                            b = b + 1
                        End If
                    End If
                    
                tot_biaya = tot_biaya + CDbl(biaya)
                    
                arr_daftar(a, 0) = b
                arr_daftar(a, 1) = tgl
                arr_daftar(a, 2) = no_bukti
                arr_daftar(a, 3) = ket
                arr_daftar(a, 4) = Format(biaya, "###,###,###")
                arr_daftar(a, 5) = user
             a = a + 1
             rs.MoveNext
             Loop
             
             grd_daftar.Columns(3).FooterText = "TOTAL"
             grd_daftar.Columns(4).FooterText = Format(tot_biaya, "###,###,###")
             
             grd_daftar.ReBind
             grd_daftar.Refresh
                    
End Sub

Private Sub cmd_cetak_Click()
On Error GoTo er_printer

    With grd_daftar.PrintInfo
        
        .PageHeaderFont.Bold = True
        .PageHeaderFont.Italic = True
        .PageHeaderFont.Size = 10
        .PageHeader = "Laporan Data Biaya"
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

On Error GoTo er_handler

Dim sql, sql1 As String
Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
    
    If arr_daftar.UpperBound(1) = 0 Then
        Exit Sub
    End If
    
    If MsgBox("Yakin akan hapus No Bukti " & no_buk, vbYesNo + vbQuestion, "Pesan") = vbNo Then
        Exit Sub
    End If
    
    sql = "select no_bukti from tbl_biaya where no_bukti='" & no_buk & "'"
    rs.Open sql, cn
        If Not rs.EOF Then
            
            sql1 = "delete from tbl_biaya where no_bukti='" & no_buk & "'"
            rs1.Open sql1, cn
            
            sql1 = "delete from tr_biaya where no_bukti='" & no_buk & "'"
            rs1.Open sql1, cn
            
         Else
            
            MsgBox ("Data yang akan dihapus tidak ditemukan")
            
         End If
    rs.Close
    cmd_tampil_Click
    Exit Sub
    
er_handler:
                
            Dim psn
                psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
                Err.Clear
End Sub

Private Sub cmd_setup_Click()
On Error GoTo er_handler
    
    With grd_daftar.PrintInfo
        .PageSetup
    End With
    Exit Sub

er_handler:
                
            Dim psn
                psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
                Err.Clear
End Sub

Private Sub cmd_tampil_Click()
    isi_daftar
End Sub


Private Sub Form_Load()
    
    grd_daftar.Array = arr_daftar
    
    kosong_daftar
    
    
    Me.Left = Screen.Width / 2 - Me.Width / 2
    Me.Top = Screen.Height / 2 - Me.Height / 2 - 2000
    
End Sub

Private Sub grd_daftar_Click()
On Error Resume Next
    If arr_daftar.UpperBound(1) > 0 Then
        no_buk = arr_daftar(grd_daftar.Bookmark, 2)
    End If
End Sub

Private Sub grd_daftar_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    grd_daftar_Click
End Sub
