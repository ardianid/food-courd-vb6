VERSION 5.00
Object = "{EC76FE26-BAFD-4E89-AA40-E748DA83A570}#1.0#0"; "IsButton_Ard.ocx"
Begin VB.Form Frm_SetPrinter 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seting Printer"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5415
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   5415
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   3615
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   1095
      Left            =   3720
      TabIndex        =   2
      Top             =   2040
      Width           =   1575
      Begin IsButton_Ard.isButton Cmd_Set 
         Height          =   735
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1296
         Icon            =   "Frm_SetPrinter.frx":0000
         Style           =   8
         Caption         =   "Set Default Printer"
         IconSize        =   32
         IconAlign       =   0
         iNonThemeStyle  =   0
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Pilih printer"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.ListBox List1 
         Height          =   1320
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   4815
      End
   End
End
Attribute VB_Name = "Frm_SetPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cSetPrinter As New cSetDfltPrinter

Private Sub cmd_set_Click()
    
    Screen.MousePointer = vbHourglass
    
    Dim sMsg As String
    Dim DeviceName As String
    
    If List1.SelCount = 1 Then
        DeviceName = List1.List(List1.ListIndex)
        If cSetPrinter.SetPrinterAsDefault(DeviceName) Then
            sMsg = DeviceName & " has successfully been set as the default printer."
        Else
            sMsg = DeviceName & " has failed to be set as the default printer."
        End If
        MsgBox sMsg, vbExclamation, App.Title
    Else
        MsgBox "Please select a printer from the list.", vbInformation, App.Title
    End If
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Load()
        
    With Me
        .Left = Screen.Width / 2 - .Width / 2
        .Top = Screen.Height / 2 - .Height / 2
    End With
        
    With Me
        .Left = Screen.Width / 2 - .Width / 2
        .Top = 250
    End With

    Dim i As Integer
    
    For i = 0 To Printers.Count - 1
        List1.AddItem Printers(i).DeviceName
    Next i


End Sub

Private Sub Form_Unload(Cancel As Integer)
    utama.Enabled = True
End Sub
