VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_browse_blg 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   10605
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      ScaleHeight     =   705
      ScaleWidth      =   10305
      TabIndex        =   12
      Top             =   5040
      Width           =   10335
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
         TabIndex        =   8
         Top             =   120
         Width           =   1335
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
         Left            =   1800
         TabIndex        =   7
         Top             =   120
         Width           =   1335
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
         Left            =   240
         TabIndex        =   6
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
         Left            =   7680
         TabIndex        =   4
         Top             =   120
         Width           =   1215
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
         Left            =   9000
         TabIndex        =   5
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   120
      ScaleHeight     =   3225
      ScaleWidth      =   10305
      TabIndex        =   11
      Top             =   1680
      Width           =   10335
      Begin TrueDBGrid60.TDBGrid grd_counter 
         Height          =   3015
         Left            =   120
         OleObjectBlob   =   "frm_browse_blg.frx":0000
         TabIndex        =   14
         Top             =   120
         Width           =   10095
      End
      Begin MSComDlg.CommonDialog cd 
         Left            =   6480
         Top             =   -120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1425
      ScaleWidth      =   10305
      TabIndex        =   0
      Top             =   120
      Width           =   10335
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
         Left            =   8880
         TabIndex        =   3
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txt_nama 
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
         Left            =   2520
         TabIndex        =   2
         Top             =   960
         Width           =   3615
      End
      Begin VB.TextBox txt_kode 
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
         Left            =   2520
         TabIndex        =   1
         Top             =   480
         Width           =   2055
      End
      Begin VB.Line Line1 
         X1              =   360
         X2              =   10080
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pencarian"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   1050
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   720
         TabIndex        =   10
         Top             =   1080
         Width           =   1515
      End
      Begin VB.Label Label1 
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
         Left            =   720
         TabIndex        =   9
         Top             =   600
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frm_browse_blg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arr_counter As New XArrayDB
Dim sql_c As String
Dim id_biling As String

Private Sub cmd_cetak_Click()
On Error GoTo er_printer

    With grd_counter.PrintInfo
        
        .PageHeaderFont.Bold = True
        .PageHeaderFont.Italic = True
        .PageHeaderFont.Size = 10
        .PageHeader = "Laporan Data Absen Billing Listrik & Air"
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
        grd_counter.Columns(4).Locked = False
        grd_counter.Columns(5).Locked = False
        grd_counter.MoveFirst
        Exit Sub
    End If
    
    If cmd_edit.Caption = "Read Only" Then
        cmd_edit.Caption = "Edit"
        grd_counter.Columns(4).Locked = True
        grd_counter.Columns(5).Locked = True
        cmd_tampil_Click
    End If
    
End Sub

Private Sub cmd_export_Click()
    
    On Error Resume Next

    cd.ShowSave
    grd_counter.ExportToFile cd.FileName, False
    
End Sub

Private Sub cmd_hapus_Click()
    Dim sql, sql1 As String
    Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
        
    On Error GoTo er_hapus
        
        If arr_counter.UpperBound(1) = 0 Then
            Exit Sub
        End If
        
        If MsgBox("Yakin akan hapus data biling counter " & arr_counter(grd_counter.Bookmark, 3), vbYesNo + vbQuestion, "Pesan") = vbNo Then
            Exit Sub
        End If
        
        sql = "select id from tbl_biling where id=" & id_biling
        rs.Open sql, cn
            If Not rs.EOF Then
                
                sql1 = "delete from tbl_biling where id=" & id_biling
                rs1.Open sql1, cn
                
            Else
                
                MsgBox ("Data yang akan dihapus tidak ditemukan")
                
            End If
        rs.Close
        cmd_tampil_Click
        Exit Sub
        
er_hapus:
        Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
        
End Sub

Private Sub cmd_setup_Click()
On Error GoTo er_set
        
        With grd_counter.PrintInfo
            .PageSetup
        End With
        Exit Sub
        
er_set:
        Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
End Sub

Private Sub cmd_tampil_Click()
    isi_counter
End Sub

Private Sub Form_Load()

    grd_counter.Array = arr_counter
    
    kosong_counter
    
    Call cari_wewenang("Form Data Billing")
      
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
    
    Me.Left = utama.Width / 2 - Me.Width / 2
    Me.Top = utama.Height / 2 - Me.Height / 2 - 2300
    
End Sub

Private Sub kosong_counter()
    arr_counter.ReDim 0, 0, 0, 0
    grd_counter.ReBind
    grd_counter.Refresh
End Sub

Private Sub isi_counter()
    
    On Error GoTo er_counter
    
    Dim rs As New ADODB.Recordset
        
        kosong_counter
        
        sql_c = "select * from qr_biling_aja"
        
        If txt_kode.Text <> "" Or txt_nama.Text <> "" Then
            sql_c = sql_c & " where"
                
                If txt_kode.Text <> "" Then
                    sql_c = sql_c & " kode like '%" & Trim(txt_kode.Text) & "%'"
                End If
                
                If txt_nama.Text <> "" And txt_kode.Text = "" Then
                    sql_c = sql_c & " nama_counter like '%" & Trim(txt_nama.Text) & "%'"
                End If
                
                If txt_nama.Text <> "" And txt_kode.Text <> "" Then
                    sql_c = sql_c & " and nama_counter like '%" & Trim(txt_nama.Text) & "%'"
                End If
        End If
        
        rs.Open sql_c, cn, adOpenKeyset
            If Not rs.EOF Then
                
                rs.MoveLast
                rs.MoveFirst
                    
                lanjut_isi rs
            End If
        rs.Close
        
        Exit Sub
        
er_counter:
        Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
        
End Sub

Private Sub lanjut_isi(rs As Recordset)
    Dim idn, kd, nm, lstr, ar As String
    Dim a As Long
        
        a = 1
        Do While Not rs.EOF
            arr_counter.ReDim 1, a, 0, 7
            grd_counter.ReBind
            grd_counter.Refresh
                
                idn = rs("id")
                
                If Not IsNull(rs("kode")) Then
                    kd = rs("kode")
                Else
                    kd = ""
                End If
                
                If Not IsNull(rs("nama_counter")) Then
                    nm = rs("nama_counter")
                Else
                    nm = ""
                End If
                
                If Not IsNull(rs("harga_listrik")) Then
                    lstr = rs("harga_listrik")
                Else
                    lstr = 0
                End If
                
                If Not IsNull(rs("harga_air")) Then
                    ar = rs("harga_air")
                Else
                    ar = 0
                End If
                
            arr_counter(a, 0) = idn
            arr_counter(a, 1) = a
            arr_counter(a, 2) = kd
            arr_counter(a, 3) = nm
            arr_counter(a, 4) = lstr
            arr_counter(a, 5) = ar
            
        a = a + 1
        rs.MoveNext
        Loop
            grd_counter.ReBind
            grd_counter.Refresh
End Sub
    
Private Sub grd_counter_AfterColUpdate(ByVal ColIndex As Integer)

Dim sql As String
Dim rs As New ADODB.Recordset

On Error GoTo er_e

    If ColIndex = 4 Then
        arr_counter(grd_counter.Bookmark, ColIndex) = grd_counter.Columns(ColIndex).Text
                
        sql = "update tbl_biling set harga_listrik=" & CCur(arr_counter(grd_counter.Bookmark, ColIndex)) & " where id=" & id_biling
        rs.Open sql, cn
        Exit Sub
    End If
        
    If ColIndex = 5 Then
        arr_counter(grd_counter.Bookmark, ColIndex) = grd_counter.Columns(ColIndex).Text
                
        sql = "update tbl_biling set harga_air=" & CCur(arr_counter(grd_counter.Bookmark, ColIndex)) & " where id=" & id_biling
        rs.Open sql, cn
        Exit Sub
    End If
        
er_e:
        Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
        
End Sub

Private Sub grd_counter_Click()
    On Error Resume Next
        If arr_counter.UpperBound(1) > 0 Then
            id_biling = arr_counter(grd_counter.Bookmark, 0)
        End If
End Sub

Private Sub grd_counter_HeadClick(ByVal ColIndex As Integer)

On Error GoTo er_head

    Dim sql As String
    Dim rs As New ADODB.Recordset
    
    If sql_c = "" Then
        Exit Sub
    End If
    
    If arr_counter.UpperBound(1) = 0 Then
        Exit Sub
    End If
    
    sql = ""
    sql = sql & sql_c
    
    Select Case ColIndex
        Case 2
            sql = sql & " order by kode"
        Case 3
            sql = sql & " order by nama_counter"
        Case 4
            sql = sql & " order by harga_listrik"
        Case 5
            sql = sql & " order by harga_air"
    End Select
                
    rs.Open sql, cn, adOpenKeyset
        If Not rs.EOF Then
            
            rs.MoveLast
            rs.MoveFirst
            
            lanjut_isi rs
        End If
    rs.Close
    
    Exit Sub
    
er_head:
    Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
    
End Sub


Private Sub grd_counter_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    grd_counter_Click
End Sub

Private Sub txt_kode_GotFocus()
    txt_kode.SelStart = 0
    txt_kode.SelLength = Len(txt_kode)
End Sub
Private Sub txt_nama_GotFocus()
    txt_nama.SelStart = 0
    txt_nama.SelLength = Len(txt_nama)
End Sub
