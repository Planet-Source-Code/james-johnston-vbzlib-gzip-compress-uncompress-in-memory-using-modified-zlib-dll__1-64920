VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmMain 
   Caption         =   "VBZLib"
   ClientHeight    =   7665
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   ScaleHeight     =   7665
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   465
      Left            =   5895
      TabIndex        =   7
      Top             =   135
      Width           =   1185
   End
   Begin VB.CommandButton cmdUncompressZLIB 
      Caption         =   "Uncompress ZLIB"
      Height          =   465
      Left            =   5895
      TabIndex        =   6
      Top             =   5190
      Width           =   1185
   End
   Begin VB.CommandButton cmdCompressZLIB 
      Caption         =   "Compress ZLIB"
      Height          =   465
      Left            =   5895
      TabIndex        =   5
      Top             =   2670
      Width           =   1185
   End
   Begin VB.CommandButton cmdUncompressGZ 
      Caption         =   "Uncompress GZ"
      Height          =   465
      Left            =   5895
      TabIndex        =   3
      Top             =   5730
      Width           =   1185
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   465
      Left            =   5895
      TabIndex        =   2
      Top             =   675
      Width           =   1185
   End
   Begin VB.CommandButton cmdCompressGZ 
      Caption         =   "Compress    GZ"
      Height          =   465
      Left            =   5895
      TabIndex        =   1
      Top             =   3210
      Width           =   1185
   End
   Begin VB.CommandButton cmdUncompressAuto 
      Caption         =   "Uncompress Auto"
      Height          =   465
      Left            =   5895
      TabIndex        =   0
      Top             =   6270
      Width           =   1185
   End
   Begin RichTextLib.RichTextBox rtbInput 
      Height          =   2445
      Left            =   90
      TabIndex        =   4
      Top             =   90
      Width           =   5730
      _ExtentX        =   10107
      _ExtentY        =   4313
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmMain.frx":0000
   End
   Begin RichTextLib.RichTextBox rtbCompress 
      Height          =   2445
      Left            =   90
      TabIndex        =   8
      Top             =   2610
      Width           =   5730
      _ExtentX        =   10107
      _ExtentY        =   4313
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmMain.frx":0082
   End
   Begin RichTextLib.RichTextBox rtbUncompress 
      Height          =   2445
      Left            =   90
      TabIndex        =   9
      Top             =   5130
      Width           =   5730
      _ExtentX        =   10107
      _ExtentY        =   4313
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmMain.frx":0104
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oZLIB As New cZLIB

'Load Sample Text
Private Sub cmdLoad_Click()
    Dim iInFile As Long, sInput As String
    sInput = LoadData(App.Path & "\google.txt")
    rtbInput.Text = sInput
End Sub

'Clear all Text
Private Sub cmdReset_Click()
    rtbInput.Text = vbNullString
    rtbCompress.Text = vbNullString
    rtbUncompress.Text = vbNullString
End Sub

'Compress with ZLIB Header
Private Sub cmdCompressZLIB_Click()
    Dim sData As String
    
    sData = rtbInput.Text
    Call oZLIB.CompressString(sData, Z_ZLIB)
    rtbCompress.Text = sData
    SaveData App.Path & "\google.z", sData
End Sub

'Compress with GZip Header
Private Sub cmdCompressGZ_Click()
    Dim sData As String
    
    sData = rtbInput.Text
    Call oZLIB.CompressString(sData, Z_GZIP)
    rtbCompress.Text = sData
    SaveData App.Path & "\google.gz", sData
End Sub


'Uncompress ZLIB data
Private Sub cmdUncompressZLIB_Click()
    Dim sData As String
    
    sData = rtbInput.Text
    Call oZLIB.UncompressString(sData, Z_ZLIB)
    rtbUncompress.Text = sData
End Sub

'Uncompress GZip data
Private Sub cmdUncompressGZ_Click()
    Dim sData As String
    
    sData = rtbInput.Text
    Call oZLIB.UncompressString(sData, Z_GZIP)
    rtbUncompress.Text = sData
End Sub

'Uncompress and auto detect data header
Private Sub cmdUncompressAuto_Click()
    Dim sData As String
    
    sData = rtbInput.Text
    Call oZLIB.UncompressString(sData, Z_AUTO)
    rtbUncompress.Text = sData
End Sub


'Load Text data
Private Function LoadData(Filename As String) As String
    Dim iInFile As Long, sInput As String
    
    iInFile = FreeFile
    Open Filename For Binary As #iInFile
    sInput = Space(LOF(iInFile))
    Get #iInFile, , sInput
    Close #iInFile
    
    LoadData = sInput
End Function

'Save Text data
Private Function SaveData(Filename As String, Data As String)
    Dim iInFile As Long
    
    iInFile = FreeFile
    Open Filename For Binary As #iInFile
    Put #iInFile, , Data
    Close #iInFile
End Function


