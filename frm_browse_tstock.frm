VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_browse_tstock 
   ClientHeight    =   8385
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8385
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   8175
      Left            =   360
      ScaleHeight     =   8145
      ScaleWidth      =   14505
      TabIndex        =   0
      Top             =   120
      Width           =   14535
      Begin VB.CommandButton cmd_setup 
         Caption         =   "Page Setup"
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
         Left            =   9840
         TabIndex        =   15
         Top             =   7560
         Width           =   1455
      End
      Begin VB.CommandButton cmd_cetak 
         Caption         =   "Cetak"
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
         Left            =   11400
         TabIndex        =   14
         Top             =   7560
         Width           =   1455
      End
      Begin MSComDlg.CommonDialog cd 
         Left            =   1800
         Top             =   3840
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmd_export 
         Caption         =   "Export"
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
         Left            =   12960
         TabIndex        =   13
         Top             =   7560
         Width           =   1455
      End
      Begin TrueDBGrid60.TDBGrid grd_daftar 
         Height          =   5655
         Left            =   120
         OleObjectBlob   =   "frm_browse_tstock.frx":0000
         TabIndex        =   12
         Top             =   1800
         Width           =   14295
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   120
         ScaleHeight     =   1545
         ScaleWidth      =   14265
         TabIndex        =   1
         Top             =   120
         Width           =   14295
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
            Height          =   495
            Left            =   12840
            TabIndex        =   11
            Top             =   840
            Width           =   1335
         End
         Begin VB.TextBox txt_nama_barang 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6360
            TabIndex        =   10
            Top             =   960
            Width           =   3495
         End
         Begin VB.TextBox txt_nama_counter 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6360
            TabIndex        =   9
            Top             =   480
            Width           =   3495
         End
         Begin VB.TextBox txt_kode_barang 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2520
            TabIndex        =   8
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox txt_kode_counter 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2520
            TabIndex        =   7
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Barang"
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
            Left            =   4080
            TabIndex        =   6
            Top             =   960
            Width           =   1350
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Counter"
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
            Left            =   4080
            TabIndex        =   5
            Top             =   480
            Width           =   1425
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Counter"
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
            Left            =   360
            TabIndex        =   4
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Barang"
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
            Left            =   360
            TabIndex        =   3
            Top             =   960
            Width           =   1260
         End
         Begin VB.Line Line1 
            X1              =   240
            X2              =   14160
            Y1              =   360
            Y2              =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pencarian"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   330
            Left            =   120
            TabIndex        =   2
            Top             =   0
            Width           =   1485
         End
      End
   End
End
Attribute VB_Name = "frm_browse_tstock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arr_daftar As New XArrayDB
Dim sql As String, id_st, id_st1 As String

Private Sub cmd_cetak_Click()
    On Error GoTo er_printer

    With grd_daftar.PrintInfo
        
        .PageHeaderFont.Bold = True
        .PageHeaderFont.Italic = True
        .PageHeaderFont.Size = 10
        .PageHeader = "Laporan Persediaan Stock Barang"
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
    
    On Error GoTo er_s
    
    With grd_daftar.PrintInfo
        .PageSetup
    End With
    Exit Sub
    
er_s:
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear

End Sub

Private Sub Cmd_Tampil_Click()

On Error GoTo er_tampil

    Dim rs As New ADODB.Recordset
        
        kosong_daftar
        
        sql = "select * from qr_stock"
        
     If txt_kode_counter.Text <> "" Or txt_kode_barang.Text <> "" Or txt_nama_barang.Text <> "" Or txt_nama_counter.Text <> "" Then
        sql = sql & " Where"
        
        If txt_kode_counter.Text <> "" Then
            sql = sql & " kode_counter like '%" & Trim(txt_kode_counter.Text) & "%'"
        End If
        
        If txt_kode_barang.Text <> "" And txt_kode_counter.Text = "" Then
            sql = sql & " kode_barang like '%" & Trim(txt_kode_barang.Text) & "%'"
        End If
        
        If txt_kode_barang.Text <> "" And txt_kode_counter.Text <> "" Then
            sql = sql & " and kode_barang like '%" & Trim(txt_kode_barang.Text) & "%'"
        End If
        
        If txt_nama_counter.Text <> "" And txt_kode_counter.Text = "" And txt_kode_barang.Text = "" Then
            sql = sql & " nama_counter like '%" & Trim(txt_nama_counter.Text) & "%'"
        End If
        
        If txt_nama_counter.Text <> "" And (txt_kode_counter.Text <> "" Or txt_kode_barang.Text <> "") Then
            sql = sql & " and nama_counter like '%" & Trim(txt_nama_counter.Text) & "%'"
        End If
        
        If txt_nama_barang.Text <> "" And txt_kode_counter.Text = "" And txt_kode_barang.Text = "" And txt_nama_counter.Text = "" Then
            sql = sql & " nama_barang like '%" & Trim(txt_nama_barang.Text) & "%'"
        End If
        
        If txt_nama_barang.Text <> "" And (txt_kode_counter.Text <> "" Or txt_kode_barang.Text <> "" Or txt_nama_counter.Text <> "") Then
            sql = sql & " and nama_barang like '%" & Trim(txt_nama_barang.Text) & "%'"
        End If
     End If
     
     sql = sql & " order by kode_counter"
     rs.Open sql, cn, adOpenKeyset
        If Not rs.EOF Then
            
            rs.MoveLast
            rs.MoveFirst
            
            isi_daftar rs
            
        End If
     rs.Close
     
     Exit Sub
     
er_tampil:
        Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
     
End Sub

Private Sub isi_daftar(rs As Recordset)
    Dim stock_min, stock_max, kode_counter, nama_counter, kode_barang, nama_barang, stock As String
    Dim a, b As Long
        
        a = 1
        b = 1
            Do While Not rs.EOF
                arr_daftar.ReDim 1, a, 0, 9
                grd_daftar.ReBind
                grd_daftar.Refresh
                       
                  
                       
                  If Not IsNull(rs("kode_counter")) Then
                    kode_counter = rs("kode_counter")
                  Else
                    kode_counter = ""
                  End If
                  
                  If Not IsNull(rs("nama_counter")) Then
                    nama_counter = rs("nama_counter")
                  Else
                    nama_counter = ""
                  End If
                  
                  If Not IsNull(rs("kode_barang")) Then
                    kode_barang = rs("kode_barang")
                  Else
                    kode_barang = ""
                  End If
                  
                  If Not IsNull(rs("nama_barang")) Then
                    nama_barang = rs("nama_barang")
                  Else
                    nama_barang = ""
                  End If
                  
                  If Not IsNull(rs("stock_min")) Then
                    stock_min = rs("stock_min")
                  Else
                    stock_min = 0
                  End If
                  
                  If Not IsNull(rs("stock_max")) Then
                    stock_max = rs("stock_max")
                  Else
                    stock_max = 0
                  End If
                  
                  If Not IsNull(rs("jml_stock")) Then
                    stock = rs("jml_stock")
                  Else
                    stock = 0
                  End If
                           
                If a > 1 Then
                    If kode_counter <> "" And (kode_counter <> arr_daftar(a - 1, 1)) Then
                        b = b + 1
                    End If
                End If
                           
                arr_daftar(a, 0) = b
                arr_daftar(a, 1) = kode_counter
                arr_daftar(a, 2) = nama_counter
                arr_daftar(a, 3) = kode_barang
                arr_daftar(a, 4) = nama_barang
                arr_daftar(a, 5) = stock_min
                arr_daftar(a, 6) = stock_max
                arr_daftar(a, 7) = stock
            
            a = a + 1
            rs.MoveNext
            Loop
            grd_daftar.ReBind
            grd_daftar.Refresh
        
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


Private Sub grd_daftar_HeadClick(ByVal ColIndex As Integer)

On Error GoTo er_head

    Dim sql1 As String
    Dim rs As New ADODB.Recordset
        
    sql1 = ""
        
        If sql = "" Then
            Exit Sub
        End If
        
If arr_daftar.UpperBound(1) > 0 Then
    sql1 = sql
        
        Select Case ColIndex
            Case 2
                sql1 = sql1 & ",nama_counter"
            Case 3
                sql1 = sql1 & ",kode_barang"
            Case 4
                sql1 = sql1 & ",nama_barang"
            Case 5
                sql1 = sql1 & ",stock_min"
            Case 6
                sql1 = sql1 & ",stock_max"
            Case 7
                sql1 = sql1 & ",jml_stock"
        End Select
        
        
    rs.Open sql, cn, adOpenKeyset
        If Not rs.EOF Then
            
            rs.MoveLast
            rs.MoveFirst
                
                isi_daftar rs
        End If
    rs.Close
End If
        
Exit Sub

er_head:
    Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
        
End Sub
        

