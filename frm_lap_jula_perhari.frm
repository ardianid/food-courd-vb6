VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_lap_jual_perhari 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Laporan Per Counter"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   120
      ScaleHeight     =   2025
      ScaleWidth      =   6345
      TabIndex        =   0
      Top             =   120
      Width           =   6375
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
         Left            =   5160
         TabIndex        =   5
         Top             =   1440
         Width           =   1095
      End
      Begin Crystal.CrystalReport rpt 
         Left            =   120
         Top             =   1680
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowTitle     =   "LAPORAN PENJUALAN PERHARI"
         WindowBorderStyle=   0
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowGroupTree=   -1  'True
         WindowAllowDrillDown=   -1  'True
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin VB.TextBox txt_kode 
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
         Height          =   375
         Left            =   2160
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
      Begin MSMask.MaskEdBox msk_tgl 
         Height          =   375
         Left            =   2160
         TabIndex        =   2
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
      Begin MSMask.MaskEdBox msk_tgl2 
         Height          =   375
         Left            =   4440
         TabIndex        =   3
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "s/d"
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
         Left            =   3960
         TabIndex        =   7
         Top             =   240
         Width           =   315
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kd Counter"
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
         Left            =   825
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl"
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
         Left            =   1560
         TabIndex        =   1
         Top             =   240
         Width           =   300
      End
   End
End
Attribute VB_Name = "frm_lap_jual_perhari"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sql As String
Dim rs As New ADODB.Recordset

Private Sub Cmd_Tampil_Click()
    
    On Error GoTo er_tampil
    
    Me.MousePointer = vbHourglass
    utama.MousePointer = vbHourglass
    
    If rs.State = adStateOpen Then
        rs.Close
    End If
   
   tgl1 = msk_tgl.Text
   tgl2 = msk_tgl2.Text
   
        sql = "select * From qr_penjualan_sebenarnya"
   
    If (msk_tgl.Text <> "__/__/____" And msk_tgl2.Text <> "__/__/____") Or txt_kode.Text <> "" Then
        
        sql = sql & " where"
            
            If msk_tgl.Text <> "__/__/____" And msk_tgl2.Text <> "__/__/____" Then
                
                sql = sql & " tgl >= datevalue('" & Trim(msk_tgl.Text) & "') and tgl <= datevalue('" & Trim(msk_tgl2.Text) & "')"
                
            End If
            
            If txt_kode.Text <> "" And (msk_tgl.Text = "__/__/____" And msk_tgl2.Text = "__/__/____") Then
                
                sql = sql & " kode_counter='" & Trim(txt_kode.Text) & "'"
                
            End If
            
            If txt_kode.Text <> "" And (msk_tgl.Text <> "__/__/____" And msk_tgl2.Text <> "__/__/____") Then
                
                sql = sql & " and kode_counter='" & Trim(txt_kode.Text) & "'"
                
            End If
            
                
    End If
    
    sqlku = sql
    
    Load frm_lap_counter
    frm_lap_counter.Show
    
      Me.MousePointer = vbDefault
    utama.MousePointer = vbDefault
    
    Exit Sub
    
er_tampil:
        
        If Me.MousePointer = vbHourglass Then
            Me.MousePointer = vbDefault
            utama.MousePointer = vbDefault
        End If
        
        Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
    
End Sub

Private Sub Form_Activate()
On Error Resume Next
    msk_tgl.SetFocus
End Sub

Private Sub Form_Load()
    
    With Me
        .Left = Screen.Width / 2 - .Width / 2
        .Top = 1000
    End With
    
End Sub

Private Sub msk_tgl_GotFocus()
    msk_tgl.SelStart = 0
    msk_tgl.SelLength = Len(msk_tgl)
End Sub

Private Sub msk_tgl2_GotFocus()
    msk_tgl.SelStart = 0
    msk_tgl.SelLength = Len(msk_tgl)
End Sub

Private Sub txt_kode_GotFocus()
    txt_kode.SelStart = 0
    txt_kode.SelLength = Len(txt_kode)
End Sub
