VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_histo_invent 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form Histori Inventori"
   ClientHeight    =   9540
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12060
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9540
   ScaleWidth      =   12060
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   9255
      Left            =   120
      ScaleHeight     =   9225
      ScaleWidth      =   11745
      TabIndex        =   0
      Top             =   120
      Width           =   11775
      Begin MSComDlg.CommonDialog cd1 
         Left            =   3000
         Top             =   6120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame Frame2 
         Height          =   855
         Left            =   5400
         TabIndex        =   12
         Top             =   8160
         Width           =   5655
         Begin VB.CommandButton cmd_keluar 
            Caption         =   "&Keluar"
            Height          =   495
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmd_Export 
            Caption         =   "Export"
            Height          =   495
            Left            =   4200
            TabIndex        =   15
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmd_cetak 
            Caption         =   "Cetak"
            Height          =   495
            Left            =   2760
            TabIndex        =   14
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmd_setup 
            Caption         =   "Page Setup"
            Height          =   495
            Left            =   1440
            TabIndex        =   13
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2655
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   11295
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   240
            ScaleHeight     =   465
            ScaleWidth      =   5145
            TabIndex        =   16
            Top             =   1920
            Width           =   5175
            Begin VB.CheckBox opt_semua 
               Caption         =   "Semua"
               Height          =   495
               Left            =   3480
               TabIndex        =   19
               Top             =   0
               Width           =   1575
            End
            Begin VB.CheckBox opt_in 
               Caption         =   "Invent In"
               Height          =   495
               Left            =   120
               TabIndex        =   18
               Top             =   0
               Width           =   1815
            End
            Begin VB.CheckBox opt_out 
               Caption         =   "Invent Out"
               Height          =   495
               Left            =   1920
               TabIndex        =   17
               Top             =   0
               Width           =   1575
            End
         End
         Begin VB.TextBox txt_nama_inventori 
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
            Height          =   405
            Left            =   1920
            TabIndex        =   5
            Top             =   1320
            Width           =   3495
         End
         Begin VB.TextBox txt_kode_inventori 
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
            Height          =   405
            Left            =   1920
            TabIndex        =   4
            Top             =   840
            Width           =   1815
         End
         Begin VB.CommandButton cmd_tampil 
            Caption         =   "&Tampil"
            Height          =   495
            Left            =   6000
            TabIndex        =   3
            Top             =   1920
            Width           =   1455
         End
         Begin MSMask.MaskEdBox msk_tgl 
            Height          =   375
            Left            =   1920
            TabIndex        =   10
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
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
            Left            =   4440
            TabIndex        =   11
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
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
            Left            =   240
            TabIndex        =   9
            Top             =   360
            Width           =   780
         End
         Begin VB.Label Label2 
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
            Top             =   360
            Width           =   330
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Inventori"
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
            Left            =   240
            TabIndex        =   7
            Top             =   840
            Width           =   1380
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Inventori"
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
            Left            =   240
            TabIndex        =   6
            Top             =   1320
            Width           =   1455
         End
      End
      Begin TrueOleDBGrid60.TDBGrid grd_invent 
         Height          =   5055
         Left            =   120
         OleObjectBlob   =   "frm_histo_invent.frx":0000
         TabIndex        =   1
         Top             =   3000
         Width           =   11295
      End
   End
End
Attribute VB_Name = "frm_histo_invent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim arr_invent As New XArrayDB

Private Sub cmd_cetak_Click()
      On Error GoTo er_printer

    With grd_invent.PrintInfo
        
        .PageHeaderFont.Bold = True
        .PageHeaderFont.Italic = True
        .PageHeaderFont.Size = 10
        .PageHeader = "Laporan Historical Inventori"
        .RepeatColumnHeaders = True
        .PageFooter = "\tPage: \p"
        .PrintPreview
    End With
    Exit Sub
    
er_printer:
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clea
End Sub

Private Sub cmd_Export_Click()
    On Error Resume Next

    cd1.ShowSave
    grd_invent.ExportToFile cd1.FileName, False
    
End Sub

Private Sub cmd_keluar_Click()
    Unload Me
End Sub

Private Sub cmd_setup_Click()
On Error GoTo er_setup
   
   With grd_invent.PrintInfo
        .PageSetup
   End With
   Exit Sub
   
er_setup:
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
End Sub

Private Sub cmd_tampil_Click()
   grd_invent.Array = arr_invent
   tmp_invent
End Sub

Private Sub Form_Load()
    Me.Height = 10050
    Me.Width = 12180
    Me.Left = (menu_utama.Width - frm_histo_invent.Width) / 2
    Me.Top = (menu_utama.Height - frm_histo_invent.Height) / 4
    
    
End Sub

Public Sub tmp_invent()
    Dim sql_tmp_invent As String
    Dim rs_tmp_invent As New ADODB.Recordset
    
    sql_tmp_invent = "select * from qr_tr_inventori "
    
    If msk_tgl.Text <> "__/__/____" Or msk_tgl1.Text <> "__/__/____" Or txt_kode_inventori <> "" Or txt_nama_inventori.Text <> "" Or opt_in.Value <> 0 Or opt_out.Value <> 0 Then
        sql_tmp_invent = sql_tmp_invent & " where"
        If msk_tgl.Text <> "__/__/____" And msk_tgl1.Text = "__/__/____" Then
           sql_tmp_invent = sql_tmp_invent & " tgl_tr= DateValue('" & Trim(msk_tgl.Text) & "')"
        End If
        
        If msk_tgl.Text = "__/__/____" And msk_tgl1.Text <> "__/__/____" Then
            sql_tmp_invent = sql_tmp_invent & " tgl_tr=DateValue('" & Trim(msk_tgl1.Text) & "')"
        End If
        
        If msk_tgl.Text <> "__/__/____" And msk_tgl1.Text <> "__/__/____" Then
            sql_tmp_invent = sql_tmp_invent & " tgl_tr>= DateValue('" & Trim(msk_tgl.Text) & "') and tgl_tr<=DateValue('" & Trim(msk_tgl1.Text) & "')"
        End If
        
        If txt_kode_inventori <> "" And msk_tgl.Text <> "__/__/____" Or msk_tgl1.Text <> "__/__/____" Then
            sql_tmp_invent = sql_tmp_invent & " and kode_invent like '%" & Trim(txt_kode_inventori.Text) & "%'"
        End If
        
        If txt_kode_inventori <> "" And msk_tgl.Text = "__/__/____" And msk_tgl1.Text = "__/__/____" Then
            sql_tmp_invent = sql_tmp_invent & " kode_invent like '%" & Trim(txt_kode_inventori.Text) & "%'"
        End If
        
        If txt_nama_inventori.Text <> "" And msk_tgl.Text <> "__/__/____" Or msk_tgl1.Text <> "__/__/____" Or txt_kode_inventori.Text <> "" Then
            sql_tmp_invent = sql_tmp_invent & " and nama_invent like  '%" & Trim(txt_nama_inventori.Text) & "%'"
        End If

        If txt_nama_inventori.Text <> "" And txt_kode_inventori = "" And msk_tgl.Text = "__/__/____" And msk_tgl1.Text = "__/__/____" Then
            sql_tmp_invent = sql_tmp_invent & " nama_invent like '%" & Trim(txt_nama_inventori.Text) & "%' "
        End If
        
        If opt_in.Value = 1 And txt_nama_inventori.Text <> "" Or msk_tgl.Text <> "__/__/____" Or msk_tgl1.Text <> "__/__/____" Or txt_kode_inventori.Text <> "" Then
            sql_tmp_invent = sql_tmp_invent & " and invent_in <> 0 "
        End If

        If opt_in.Value = 1 And txt_nama_inventori.Text = "" And txt_kode_inventori = "" And msk_tgl.Text = "__/__/____" And msk_tgl1.Text = "__/__/____" Then
            sql_tmp_invent = sql_tmp_invent & " invent_in <> 0 "
        End If

        If opt_out.Value <> 0 And txt_nama_inventori.Text <> "" Or msk_tgl.Text <> "__/__/____" Or msk_tgl1.Text <> "__/__/____" Or txt_kode_inventori.Text <> "" Then
            sql_tmp_invent = sql_tmp_invent & " and invent_out <> 0 "
        End If

        If opt_out.Value <> 0 And txt_nama_inventori.Text = "" And txt_kode_inventori = "" And msk_tgl.Text = "__/__/____" And msk_tgl1.Text = "__/__/____" Then
            sql_tmp_invent = sql_tmp_invent & " invent_out <> 0 "
        End If
       Else
        If opt_semua.Value <> 0 Then
            sql_tmp_invent = sql_tmp_invent
        End If
    End If
    
    
                       
    rs_tmp_invent.Open sql_tmp_invent, cn, adOpenKeyset
    
    If Not rs_tmp_invent.EOF Then
        rs_tmp_invent.MoveLast
        rs_tmp_invent.MoveFirst
        
        next_invent rs_tmp_invent
    End If
    rs_tmp_invent.Close
        
End Sub

Private Sub next_invent(rs_invent As Recordset)
    Dim tgl_tr, kode_invent, nama_invent, invent_in, invent_out, ket, id_user As String
    Dim a As Long
            
            a = 1
                Do While Not rs_invent.EOF
                    arr_invent.ReDim 1, a, 0, 6
                    grd_invent.ReBind
                    grd_invent.Refresh
                        
                        If Not IsNull(rs_invent("tgl_tr")) Then
                            tgl_tr = rs_invent("tgl_tr")
                        Else
                            tgl_tr = ""
                        End If
                        
                        If Not IsNull(rs_invent("kode_invent")) Then
                            kode_invent = rs_invent("kode_invent")
                        Else
                            kode_invent = ""
                        End If
                        
                        If Not IsNull(rs_invent("nama_invent")) Then
                            nama_invent = rs_invent("nama_invent")
                        Else
                            nama_invent = ""
                        End If
                        
                        If Not IsNull(rs_invent("invent_in")) Then
                            invent_in = rs_invent("invent_in")
                        Else
                            invent_in = ""
                        End If
                        
                        If Not IsNull(rs_invent("invent_out")) Then
                            invent_out = rs_invent("invent_out")
                        Else
                            invent_out = ""
                        End If
                        
                        If Not IsNull(rs_invent("ket")) Then
                            ket = rs_invent("ket")
                        Else
                            ket = ""
                        End If
                        
                        If Not IsNull(rs_invent("id_user")) Then
                            user_id = rs_invent("id_user")
                        Else
                            user_id = ""
                        End If
                        
                     arr_invent(a, 0) = tgl_tr
                     arr_invent(a, 1) = kode_invent
                     arr_invent(a, 2) = nama_invent
                     arr_invent(a, 3) = invent_in
                     arr_invent(a, 4) = invent_out
                     arr_invent(a, 5) = ket
                     arr_invent(a, 6) = user_id
                     
                     a = a + 1
                     rs_invent.MoveNext
                     Loop
                     grd_invent.ReBind
                     grd_invent.Refresh
End Sub


Private Sub opt_in_Click()
    If opt_in.Value = 1 Then
        opt_out.Value = 0
        'opt_semua.Value = 0
    End If
        
End Sub

Private Sub opt_out_Click()
    If opt_out.Value = 1 Then
        opt_in.Value = 0
        'opt_semua.Value = 0
    End If
    
End Sub

Private Sub opt_semua_Click()
    If opt_semua.Value = 1 Then
        opt_in.Value = 0
        opt_out.Value = 0
    End If
End Sub
