VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Without Common Dialog"
   ClientHeight    =   3924
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   2844
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3924
   ScaleWidth      =   2844
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSaveAs 
      Caption         =   "Save As"
      Height          =   372
      Left            =   120
      TabIndex        =   9
      Top             =   3480
      Width           =   1212
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   372
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   2652
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open File"
      Height          =   852
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   1212
   End
   Begin VB.TextBox txtFolder 
      Height          =   372
      Left            =   1440
      TabIndex        =   5
      Top             =   1080
      Width           =   1332
   End
   Begin VB.CommandButton cmdFolder 
      Caption         =   "Choose Folder"
      Height          =   372
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1212
   End
   Begin VB.TextBox txtFont 
      Height          =   372
      Left            =   1440
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Text            =   "Main.frx":0000
      Top             =   600
      Width           =   1332
   End
   Begin VB.CommandButton cmdFont 
      Caption         =   "Font"
      Height          =   372
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1212
   End
   Begin VB.PictureBox picColor 
      Height          =   372
      Left            =   1440
      ScaleHeight     =   324
      ScaleWidth      =   1284
      TabIndex        =   1
      Top             =   120
      Width           =   1332
   End
   Begin VB.CommandButton cmdColor 
      Caption         =   "Color"
      Height          =   372
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1212
   End
   Begin VB.CommandButton cmdShutdown 
      Caption         =   "ShutDown dialog"
      Height          =   372
      Left            =   120
      TabIndex        =   11
      Top             =   3000
      Width           =   2652
   End
   Begin VB.Label lblSaveAs 
      Height          =   372
      Left            =   1440
      TabIndex        =   10
      Top             =   3480
      Width           =   1332
   End
   Begin VB.OLE OLE 
      Height          =   852
      Left            =   1440
      SizeMode        =   1  'Stretch
      TabIndex        =   7
      Top             =   1560
      Width           =   1332
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdColor_Click()
Dim iReturn As Long
iReturn = ShowColorDlg(Me.hWnd, picColor.BackColor, False)
If iReturn <> -1 Then
picColor.BackColor = iReturn
End If
End Sub

Private Sub cmdFolder_Click()
txtFolder.Text = CDModule.BrowseForFolder(Me.hWnd, "Choose Folder:")
End Sub

Private Sub cmdFont_Click()
CDModule.ChooseFontDialog Me.hWnd, txtFont
End Sub

Private Sub cmdOpen_Click()
Dim Filter As String
Dim Ret As String
Filter = "All files" & Chr(0) & "*.*"
'The Filter fo the dialog
Ret = CDModule.ShowOpenDlg(Me.hWnd, Filter)
If Ret <> "Cancel" Then
    On Error Resume Next
    OLE.CreateLink Ret
End If
End Sub

Private Sub cmdPrint_Click()
Dim Ret As Boolean
Ret = CDModule.ShowPrint(Me.hWnd, Me.hdc, False, 1, 10, 10, 3)
If Ret = True Then MsgBox "Successful!", vbOKOnly, "Print"
End Sub

Private Sub cmdSaveAs_Click()
Dim Ret As String
Ret = CDModule.ShowSavedlg(Me.hWnd)
If Ret <> "Cancel" Then
lblSaveAs.Caption = Ret
End If
End Sub

Private Sub cmdShutdown_Click()
Call CDModule.SHShutDownDialog(0)
End Sub
