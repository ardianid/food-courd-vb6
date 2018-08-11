VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frm_input_biaya 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   9615
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7695
      Left            =   120
      ScaleHeight     =   7665
      ScaleWidth      =   9345
      TabIndex        =   0
      Top             =   120
      Width           =   9375
      Begin MSComDlg.CommonDialog cd 
         Left            =   3240
         Top             =   5280
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame Frame1 
         Height          =   855
         Left            =   240
         TabIndex        =   10
         Top             =   6720
         Width           =   8895
         Begin VB.CommandButton cmd_et 
            Caption         =   "&Page Setup "
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
            Left            =   4560
            TabIndex        =   16
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmd_export 
            Caption         =   "&Export"
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
            Left            =   7440
            TabIndex        =   15
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmd_tambah 
            Caption         =   "&Tambah"
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
            TabIndex        =   12
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmd_baru 
            Caption         =   "&Baru "
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
            Left            =   3000
            TabIndex        =   11
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmd_cetak 
            Caption         =   "&Cetak"
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
            Left            =   6000
            TabIndex        =   14
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmd_simpan 
            Caption         =   "&Simpan"
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
            Left            =   1560
            TabIndex        =   13
            Top             =   240
            Width           =   1335
         End
      End
      Begin TrueOleDBGrid60.TDBGrid grd_biaya 
         Height          =   3975
         Left            =   240
         OleObjectBlob   =   "frm_input_biaya.frx":0000
         TabIndex        =   9
         Top             =   2640
         Width           =   8895
      End
      Begin VB.TextBox txt_jumlah 
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
         Left            =   2760
         TabIndex        =   8
         Top             =   1920
         Width           =   2655
      End
      Begin VB.TextBox txt_keterangan 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2760
         TabIndex        =   7
         Top             =   1200
         Width           =   4335
      End
      Begin VB.TextBox txt_no_bukti 
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
         Left            =   2760
         TabIndex        =   6
         Top             =   720
         Width           =   3135
      End
      Begin MSComCtl2.DTPicker dt_tanggal 
         Height          =   375
         Left            =   2760
         TabIndex        =   5
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
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
         Format          =   56098817
         CurrentDate     =   38637
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah"
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
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   2040
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan Biaya"
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
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   1725
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Bukti"
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
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
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
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   780
      End
   End
End
Attribute VB_Name = "frm_input_biaya"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql_tmp_biaya As String
Dim rs_tmp_biaya As New ADODB.Recordset
Dim konfirm As String

Private Sub cmd_baru_Click()
    Dim sql_hapus_temp As String
    
    sql_hapus_temp = "delete * from tbl_tmp_biaya "
    cn.Execute (sql_hapus_temp)
    grd_biaya.ReBind
    grd_biaya.Refresh
    
    proc_baru
    
End Sub

Private Sub cmd_cetak_Click()
On Error GoTo eror

With grd_biaya.PrintInfo
        ' Set the page header
        .PageHeaderFont.Italic = True
        .PageHeader = "Daftar Biaya Operasional"
        ' Column headers will be on every page
        .RepeatColumnHeaders = True
        ' Display page numbers (centered)
        .PageFooter = "\tPage: \p"
        ' Invoke Print Preview
        .PrintPreview
End With
    Dim sql_cetak As String

    sql_cetak = "delete from tbl_tmp_biaya"
    cn.Execute (sql_cetak)
    proc_tmp_biaya
    proc_baru

Exit Sub
     
eror:
    Dim konfirm As String
    konfirm = MsgBox("Tidak dapat menampilkan dan mencetak program" & Err.Description, vbInformation + vbOKOnly, "Informasi")
    Err.Clear
End Sub
Private Sub cmd_keluar_Click()
    Unload Me
End Sub

Private Sub cmd_et_Click()
On Error GoTo err_set
    
    With Me.grd_biaya.PrintInfo
        .PageSetup
    End With
    Exit Sub
    
err_set:
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Exit Sub
End Sub

Private Sub cmd_export_Click()
    On Error Resume Next
        cd.ShowSave
        Me.grd_biaya.ExportToFile cd.FileName, False
End Sub

Private Sub cmd_simpan_Click()

On Error GoTo er_simpan

    Dim sql_tmp_biaya As String
    Dim rs_tmp_biaya As New ADODB.Recordset
    Dim sql_smp_biaya As String
    Dim rs_smp_biaya As New ADODB.Recordset
    
    
    If txt_no_bukti.Text <> "" Then
        If Len(txt_no_bukti.Text) = 11 Then
            sql_tmp_biaya = "select * from tbl_tmp_biaya "
            Set rs_tmp_biaya = cn.Execute(sql_tmp_biaya)
            With rs_tmp_biaya
                Do While Not .EOF
                   sql_smp_biaya = "insert into tbl_biaya (tanggal,no_bukti,keterangan,biaya) values(" & _
                   "'" & Format(!tanggal, "dd/mm/yyyy") & "','" & !no_bukti & "','" & !keterangan & "','" & !biaya & "')"
                   cn.Execute (sql_smp_biaya)
                   .MoveNext
               Loop
            End With
            MsgBox ("Data berhasil disimpan")
           Else
            konfirm = MsgBox("Jumlah digit nomor Bukti harus 11, " & vbCrLf _
            & "2 digit tanggal,2 digit bulan,2 digit tahun, dan 5 digit no urut" & vbCrLf _
            & " Silahkan ganti nomor bukti anda", vbInformation + vbOKOnly, "informasi")
            txt_no_bukti.SetFocus
            Exit Sub
        End If
       Else
        konfirm = MsgBox("Input Nomor Bukti Biaya", vbInformation + vbOKOnly, "Informasi")
        txt_no_bukti.SetFocus
        Exit Sub
    End If
    Exit Sub
            
er_simpan:
    Dim pasn
        pasn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
            
End Sub

Private Sub cmd_tambah_Click()
    Dim sql_smp_biaya As String
    Dim rs_smp_biaya As New ADODB.Recordset
    
    If txt_no_bukti.Text <> "" Then
        If Len(txt_no_bukti.Text) = 11 Then
            If txt_keterangan.Text <> "" Then
                If txt_jumlah.Text <> "" Then
                   sql_smp_biaya = "insert into tbl_tmp_biaya (tanggal,no_bukti,keterangan,biaya) values(" & _
                   "'" & Format(dt_tanggal, "dd/mm/yyyy") & "','" & txt_no_bukti.Text & "','" & txt_keterangan.Text & "','" & txt_jumlah.Text & "')"
                   cn.Execute (sql_smp_biaya)
                   Else
                    konfirm = MsgBox("Input jumlah penggunaan biaya-biaya", vbInformation + vbOKOnly, "Informasi")
                    txt_jumlah.SetFocus
                    Exit Sub
                End If
               Else
                txt_keterangan.Text = "-"
            End If
           Else
            konfirm = MsgBox("Jumlah digit nomor Bukti harus 11, " & vbCrLf _
            & "2 digit tanggal,2 digit bulan,2 digit tahun, dan 5 digit no urut" & vbCrLf _
            & " Silahkan ganti nomor bukti anda", vbInformation + vbOKOnly, "informasi")
            txt_no_bukti.SetFocus
            Exit Sub
        End If
       Else
        konfirm = MsgBox("Input Nomor Bukti Biaya", vbInformation + vbOKOnly, "Informasi")
        txt_no_bukti.SetFocus
        Exit Sub
    End If
       
    proc_tmp_biaya
    proc_bersih
    
    txt_keterangan.SetFocus
End Sub

Private Sub Command1_Click()

End Sub

Private Sub dt_tanggal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        txt_no_bukti.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    Dim sql_cr_no_bukti As String
    Dim rs_cr_no_bukti As New ADODB.Recordset
    Dim no_bukti As String
    
   ' dt_tanggal.SetFocus
    proc_tmp_biaya
    proc_buat_nomor_bukti
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyF1 Then
        MsgBox ("F1")
    End If
End Sub

Private Sub Form_Load()
    dt_tanggal.Value = Format(Date, "dd/mm/yyyy")
    'Me.Height = 8880
    'Me.Width = 10065
    Me.Left = (utama.Width - frm_input_biaya.Width) / 2
    Me.Top = (utama.Height - frm_input_biaya.Height) / 2 - 1350
    
End Sub

Private Sub txt_jumlah_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If txt_jumlah.Text <> "" Then
           cmd_Tambah.SetFocus
           Else
            MsgBox ("Isis jumlah biaya")
            txt_jumlah.SetFocus
        End If
    End If
        
End Sub

Private Sub txt_jumlah_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) _
    Or KeyAscii = vbKeyBack _
    Or KeyAscii = vbKeyDelete _
    Or KeyAscii = vbKeySpace) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub txt_keterangan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If txt_keterangan.Text <> "" Then
            txt_jumlah.SetFocus
           Else
            txt_keterangan.SetFocus
        End If
    End If
    
    If KeyCode = vbKeyDown Then
        txt_jumlah.SetFocus
    End If
End Sub

Private Sub txt_keterangan_KeyPress(KeyAscii As Integer)
If Not (KeyAscii = Asc(Chr(KeyAscii))) Then
    KeyAscii = ""
End If
End Sub

Private Sub txt_no_bukti_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sql_cr_no_bukti As String
    Dim rs_no_bukti As New ADODB.Recordset
    
    If KeyCode = 13 Then
        If txt_no_bukti.Text <> "" Then
            txt_keterangan.SetFocus
          Else
           txt_no_bukti.SetFocus
        End If
    End If

        
End Sub

Public Sub proc_tmp_biaya()
    sql_tmp_biaya = "select * from tbl_tmp_biaya where no_bukti='" & txt_no_bukti.Text & "' order by tanggal"
    Set rs_tmp_biaya = cn.Execute(sql_tmp_biaya)
    Set grd_biaya.DataSource = rs_tmp_biaya
    grd_biaya.ReBind
    grd_biaya.Refresh
End Sub

Public Sub proc_bersih()
    txt_jumlah.Text = ""
    txt_keterangan.Text = ""
    txt_keterangan.SetFocus
End Sub

Public Sub proc_buat_nomor_bukti()
    On Error GoTo lanjut
    sql_cr_no_bukti = "select max(right(no_bukti,5)) as no_bukti from tbl_biaya"
    Set rs_cr_no_bukti = cn.Execute(sql_cr_no_bukti)
    With rs_cr_no_bukti
        If Not (.BOF And .EOF) Then
            no_bukti = !no_bukti
           Else
            no_bukti = "0"
        End If
        GoTo nomor
lanjut:
        no_bukti = "0"
    End With
nomor:
        no_bukti = Val(no_bukti) + 1
        no_bukti = Format(no_bukti, "00000")
    
    txt_no_bukti.Text = Format(Date, "dd") + Format(Date, "mm") + Right(Format(Date, "yyyy"), 2) + no_bukti
End Sub

Public Sub proc_baru()
    txt_jumlah.Text = ""
    txt_keterangan.Text = ""
    txt_no_bukti.Text = ""
    txt_no_bukti.Text = ""
    
    proc_buat_nomor_bukti
    txt_no_bukti.SetFocus
    proc_tmp_biaya
End Sub
