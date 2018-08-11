VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_biaya 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6375
      Left            =   120
      ScaleHeight     =   6345
      ScaleWidth      =   7785
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      Begin MSComDlg.CommonDialog cd 
         Left            =   1320
         Top             =   4320
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmd_baru 
         Caption         =   "Baru"
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
         Left            =   1440
         TabIndex        =   5
         Top             =   5640
         Width           =   1215
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
         Left            =   3240
         TabIndex        =   6
         Top             =   5640
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
         Left            =   4800
         TabIndex        =   7
         Top             =   5640
         Width           =   1335
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
         Left            =   6240
         TabIndex        =   8
         Top             =   5640
         Width           =   1335
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
         Left            =   240
         TabIndex        =   4
         Top             =   5640
         Width           =   1215
      End
      Begin TrueDBGrid60.TDBGrid grd_jumlah 
         Height          =   2775
         Left            =   240
         OleObjectBlob   =   "frm_biaya.frx":0000
         TabIndex        =   14
         Top             =   2760
         Width           =   7335
      End
      Begin VB.TextBox txt_biaya 
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
         Height          =   495
         Left            =   1800
         TabIndex        =   3
         Top             =   2040
         Width           =   2655
      End
      Begin VB.TextBox txt_ket 
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
         Height          =   495
         Left            =   1800
         TabIndex        =   2
         Top             =   1440
         Width           =   4095
      End
      Begin VB.TextBox txt_no_bukti 
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
         Height          =   495
         Left            =   1800
         TabIndex        =   1
         Top             =   840
         Width           =   2655
      End
      Begin MSComCtl2.DTPicker dtp_tgl 
         Height          =   375
         Left            =   720
         TabIndex        =   10
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   61145089
         CurrentDate     =   38639
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Biaya"
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
         Left            =   720
         TabIndex        =   13
         Top             =   2160
         Width           =   600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ket."
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
         Left            =   720
         TabIndex        =   12
         Top             =   1560
         Width           =   420
      End
      Begin VB.Label Label2 
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
         Left            =   720
         TabIndex        =   11
         Top             =   840
         Width           =   885
      End
      Begin VB.Label Label1 
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
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   375
      End
   End
End
Attribute VB_Name = "frm_biaya"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arr_jumlah As New XArrayDB
Dim angka

Private Sub kosong_jumlah()
    arr_jumlah.ReDim 0, 0, 0, 0
    grd_jumlah.ReBind
    grd_jumlah.Refresh
End Sub

Private Sub isi_bukti()

On Error GoTo er_bukti

    Dim sql, sql1 As String
    Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
    Dim bukti
        
        bukti = DatePart("d", Now)
        bukti = bukti & DatePart("m", Now)
        bukti = bukti & Right(DatePart("yyyy", Now), 2)
        bukti = bukti & "F"
        bukti = bukti & id_user
        
        sql = "select no_bukti from tbl_biaya where nama_user='" & Trim(utama.lbl_user.Caption) & "' and tgl=datevalue('" & Trim(dtp_tgl.Value) & "')"
        rs.Open sql, cn
            If Not rs.EOF Then
                sql1 = "select max(id) as mx_id from tbl_biaya where nama_user='" & Trim(utama.lbl_user.Caption) & "' and tgl=datevalue('" & Trim(dtp_tgl.Value) & "')"
                rs1.Open sql1, cn
                    If Not rs1.EOF Then
                        Dim sql2 As String
                        Dim rs2 As New ADODB.Recordset
                            sql2 = "select no_bukti from tbl_biaya where id=" & rs1("mx_id")
                            rs2.Open sql2, cn
                                If Not rs2.EOF Then
                                    Dim pecah, angka_user
                                        angka_user = Len(id_user)
                                        pecah = Mid(Trim(rs2!no_bukti), 8 + CDbl(angka_user), Len(rs2!no_bukti))
                                        pecah = CDbl(pecah) + 1
                                        
                                        bukti = bukti & pecah
                                End If
                            rs2.Close
                    End If
                rs1.Close
            Else
                bukti = bukti & 1
            End If
        rs.Close
        txt_no_bukti.Text = bukti
        
        
        Exit Sub
        
er_bukti:
        Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
End Sub

Private Sub cmd_baru_Click()
    kosong_semua
    kosong_jumlah
    grd_jumlah.Columns(1).FooterText = 0
    txt_ket.SetFocus
End Sub

Private Sub cmd_cetak_Click()
On Error GoTo er_printer

    With grd_jumlah.PrintInfo
        
        .PageHeaderFont.Bold = True
        .PageHeaderFont.Italic = True
        .PageHeaderFont.Size = 10
        .PageHeader = "Laporan Biaya    \t\tTanggal : " & dtp_tgl.Value
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
    grd_jumlah.ExportToFile cd.FileName, False
        
End Sub

Private Sub cmd_setup_Click()
    On Error GoTo aj
        
        With grd_jumlah.PrintInfo
            .PageSetup
        End With
        Exit Sub
aj:
        Dim p
            p = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
        
End Sub

Private Sub cmd_simpan_Click()

On Error GoTo er_aja

    Dim sql, sql1 As String
    Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
        
        If arr_jumlah.UpperBound(1) = 0 Then
            MsgBox ("Tidak ada data yang akan diproses")
            Exit Sub
            txt_ket.SetFocus
        End If
              
        If MsgBox("Yakin semua data yang diumasukkan sudah benar", vbYesNo + vbQuestion, "Pesan") = vbNo Then
            Exit Sub
        End If
              
    sql = "select no_bukti from tbl_biaya where no_bukti='" & Trim(txt_no_bukti.Text) & "'"
    rs.Open sql, cn
        If Not rs.EOF Then
            MsgBox ("No bukti sudah ada")
            txt_no_bukti.SetFocus
            Exit Sub
        End If
    rs.Close
    
    cn.BeginTrans
    
    sql = "insert into tbl_biaya (no_bukti,tgl,nama_user) values('" & Trim(txt_no_bukti.Text) & "','" & Trim(dtp_tgl.Value) & "','" & Trim(utama.lbl_user.Caption) & "')"
    rs.Open sql, cn
    
  Dim a As Long
    
    For a = 1 To arr_jumlah.UpperBound(1)
        sql1 = "insert into tr_biaya (no_bukti,ket,biaya) values('" & Trim(txt_no_bukti.Text) & "','" & Trim(arr_jumlah(a, 0)) & "'," & CCur(arr_jumlah(a, 1)) & ")"
        rs1.Open sql1, cn
    Next a
    
    MsgBox ("Data berhasil disimpan")
    cn.CommitTrans
    Exit Sub
    
er_aja:
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
        
End Sub

Private Sub kosong_semua()
    txt_ket.Text = ""
    txt_biaya.Text = ""
End Sub

Private Sub Form_Activate()
    txt_ket.SetFocus
End Sub

Private Sub Form_Load()

grd_jumlah.Array = arr_jumlah

kosong_jumlah

grd_jumlah.Columns(0).FooterText = "Total"
grd_jumlah.Columns(1).FooterText = 0

dtp_tgl.Value = Format(Date, "dd/mm/yyyy")

isi_bukti

Me.Left = Screen.Width / 2 - Me.Width / 2
Me.Top = Screen.Height / 2 - Me.Height / 2 - 1900

End Sub

Private Sub grd_jumlah_Click()
    On Error Resume Next
        If arr_jumlah.UpperBound(1) > 0 Then
            angka = arr_jumlah(grd_jumlah.Bookmark, 1)
        End If
End Sub

Private Sub grd_jumlah_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo er_jumlah

    If KeyCode = vbKeyDelete Then
        
        If arr_jumlah.UpperBound(1) > 1 Then
            grd_jumlah.Delete
            grd_jumlah.Columns(1).FooterText = Format(CDbl(grd_jumlah.Columns(1).FooterText) - CDbl(angka), "Currency")
        Else
            arr_jumlah.ReDim 0, 0, 0, 0
            grd_jumlah.Columns(1).FooterText = 0
        End If
            
            
            grd_jumlah.ReBind
            grd_jumlah.Refresh
    End If
        
    Exit Sub
    
er_jumlah:
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
        
End Sub

Private Sub grd_jumlah_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    grd_jumlah_Click
End Sub

Private Sub txt_biaya_GotFocus()
    txt_biaya.SelStart = 0
    txt_biaya.SelLength = Len(txt_ket)
End Sub

Private Sub txt_biaya_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo er_biaya

    If KeyCode = 13 And txt_biaya.Text <> "" Then
      arr_jumlah.ReDim 1, arr_jumlah.UpperBound(1) + 1, 0, 2
      grd_jumlah.ReBind
      grd_jumlah.Refresh
            
           grd_jumlah.Caption = ""
           grd_jumlah.Caption = "No Bukti " & Trim(txt_no_bukti.Text)
            
        Dim sk
            sk = arr_jumlah.UpperBound(1)
            
            arr_jumlah(sk, 1) = Format(txt_biaya.Text, "###,###,###")
            arr_jumlah(sk, 0) = Trim(txt_ket.Text)
            
            grd_jumlah.Columns(1).FooterText = Format(CDbl(txt_biaya.Text) + CDbl(grd_jumlah.Columns(1).FooterText), "###,###,###")
            
      
        If MsgBox("Akan menambah data lagi", vbYesNo) = vbYes Then
            kosong_semua
            txt_ket.SetFocus
            
        Else
            kosong_semua
            cmd_simpan.SetFocus
            
        End If
        
        grd_jumlah.ReBind
        grd_jumlah.Refresh
        Exit Sub
        
   End If
   
   If KeyCode = 13 And txt_biaya.Text = "" Then
        cmd_simpan.SetFocus
   End If
        
        
 Exit Sub
 
er_biaya:
        Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
End Sub

Private Sub txt_biaya_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub txt_biaya_KeyUp(KeyCode As Integer, Shift As Integer)
    If txt_biaya.Text <> "" Then
        txt_biaya.Text = Format(txt_biaya.Text, "###,###,###")
        txt_biaya.SelStart = Len(txt_biaya.Text)
    End If
End Sub

Private Sub txt_ket_GotFocus()
    txt_ket.SelStart = 0
    txt_ket.SelLength = Len(txt_ket)
End Sub

Private Sub txt_ket_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        txt_biaya.SetFocus
    End If
End Sub

Private Sub txt_no_bukti_GotFocus()
    txt_no_bukti.SelStart = 0
    txt_no_bukti.SelLength = Len(txt_no_bukti)
End Sub

Private Sub txt_no_bukti_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        txt_ket.SetFocus
    End If
End Sub
