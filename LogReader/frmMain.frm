VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log Reader"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8895
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   8895
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Slider FileSbar 
      Height          =   255
      Left            =   6720
      TabIndex        =   11
      Top             =   3720
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   1
   End
   Begin VB.CommandButton cmdReadFormHere 
      Caption         =   "Read"
      Height          =   375
      Left            =   7770
      TabIndex        =   10
      Top             =   5400
      Width           =   855
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   6720
      TabIndex        =   8
      Top             =   5400
      Width           =   855
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "Read Next"
      Height          =   375
      Left            =   7710
      TabIndex        =   7
      Top             =   4230
      Width           =   1095
   End
   Begin VB.Frame fraReadDirections 
      Caption         =   "Read Direction"
      Height          =   2535
      Left            =   6720
      TabIndex        =   4
      Top             =   600
      Width           =   2055
      Begin VB.OptionButton optEnd 
         Alignment       =   1  'Right Justify
         Caption         =   "Read From Bottom"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   2040
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton OptStart 
         Caption         =   "Read From Top"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1695
      End
      Begin VB.Image imgFile 
         Height          =   1185
         Left            =   480
         Picture         =   "frmMain.frx":058A
         Stretch         =   -1  'True
         Top             =   720
         Width           =   945
      End
      Begin VB.Line Line6 
         X1              =   1680
         X2              =   1680
         Y1              =   1560
         Y2              =   1920
      End
      Begin VB.Line Line5 
         X1              =   1800
         X2              =   1680
         Y1              =   1680
         Y2              =   1560
      End
      Begin VB.Line Line4 
         X1              =   1560
         X2              =   1680
         Y1              =   1680
         Y2              =   1560
      End
      Begin VB.Line Line3 
         X1              =   240
         X2              =   360
         Y1              =   1080
         Y2              =   960
      End
      Begin VB.Line Line2 
         X1              =   240
         X2              =   120
         Y1              =   1080
         Y2              =   960
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   240
         Y1              =   720
         Y2              =   1080
      End
   End
   Begin VB.TextBox txtFilePath 
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   240
      Width           =   5415
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Height          =   315
      Left            =   6720
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox txtBuffer 
      Height          =   3975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   600
      Width           =   6495
   End
   Begin VB.Label lblEnd 
      Caption         =   "end"
      Height          =   225
      Left            =   8550
      TabIndex        =   13
      Top             =   3480
      Width           =   285
   End
   Begin VB.Label lblStart 
      Caption         =   "start"
      Height          =   225
      Left            =   6720
      TabIndex        =   12
      Top             =   3540
      Width           =   315
   End
   Begin VB.Label Label1 
      Caption         =   "File Navigator"
      Height          =   255
      Left            =   7290
      TabIndex        =   9
      Top             =   3270
      Width           =   1005
   End
   Begin VB.Label lblFile 
      Caption         =   "File To Read"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim m_blnReadNextClicked As Boolean
Dim m_blnOpenSuccess As Boolean
Dim m_LogFile As cLogReader

Private Sub cmdBrowse_Click()
    Dim pFileName As String
    If (modMain.VBGetOpenFileName(pFileName, , True, False, True, True, , , , "Select file to read")) Then
        txtFilePath.Text = pFileName
        Call ReadyToRun
    End If
End Sub

Private Sub ReadyToRun()
    txtBuffer.Text = ""
    If (m_LogFile.FileExist(txtFilePath.Text)) Then
        Call ReInitilizeScan
    End If
    If (m_blnOpenSuccess) Then
       Call EnableControls(True)
    Else
        Call EnableControls(False)
    End If
End Sub

Private Sub cmdRead_Click()
    m_blnReadNextClicked = True
    If (txtFilePath.Text <> "") Then
        ReadFile
    End If
    m_blnReadNextClicked = False
End Sub

Private Sub EnableControls(ByVal Value As Boolean)
    Dim pBackColour As Long
    txtBuffer.Enabled = Value
    optEnd.Enabled = Value
    OptStart.Enabled = Value
    cmdRead.Enabled = Value
    cmdReadFormHere.Enabled = Value
    If (FileSbar.Tag <> "NO") Then
        FileSbar.Enabled = Value
        lblEnd.Enabled = Value
        lblStart.Enabled = Value
        Label1.Enabled = Value
    Else
        FileSbar.Enabled = False
        lblEnd.Enabled = False
        lblStart.Enabled = False
        Label1.Enabled = False
    End If
    imgFile.Enabled = Value
    fraReadDirections.Enabled = Value
    
    
    If (Value) Then
        pBackColour = RGB(0, 0, 0)
    Else
        pBackColour = RGB(100, 100, 100)
    End If
    Line1.BorderColor = pBackColour
    Line2.BorderColor = pBackColour
    Line3.BorderColor = pBackColour
    Line4.BorderColor = pBackColour
    Line5.BorderColor = pBackColour
    Line6.BorderColor = pBackColour
End Sub


Private Sub ReadFile()
If (Not m_LogFile.EOF) Then
On Error GoTo ErrorTrap

        If (FileSbar.Tag = "NO") Then
            FileSbar.Enabled = False
        Else
            FileSbar.Enabled = True
        End If
        
        If (m_LogFile.PointerPosition = 0) Then
            FileSbar.Value = 1
        ElseIf (m_LogFile.PointerPosition = m_LogFile.MaxPosition) Then
            FileSbar.Value = m_LogFile.MaxPosition
        Else
            FileSbar.Value = m_LogFile.PointerPosition
        End If
        
        If (optEnd.Value) Then
            txtBuffer.Text = m_LogFile.ReadBuffer & txtBuffer.Text
        Else
            txtBuffer.Text = txtBuffer.Text & m_LogFile.ReadBuffer
        End If
        
        
End If
Exit Sub
ErrorTrap:
    txtBuffer.Text = ""
    Resume
End Sub

Private Sub cmdReadFormHere_Click()
    txtBuffer.Text = ""
    m_LogFile.PointerPosition = FileSbar.Value
End Sub

Private Sub cmdReset_Click()
    '
    If (m_LogFile.IsFileOpened) Then m_LogFile.fClose
    
    m_LogFile.OpenLogFile txtFilePath.Text
    m_blnOpenSuccess = m_LogFile.MaxPosition
    If (m_blnOpenSuccess) Then
        m_LogFile.ScanFDirection = OptStart.Value
        If (m_LogFile.MaxPosition > 1) Then
            With FileSbar
                .Min = 1
                .Max = m_LogFile.MaxPosition
                .Value = m_LogFile.PointerPosition
            End With
        FileSbar.Tag = ""
        Else
            FileSbar.Tag = "NO"
            FileSbar.Enabled = False
            FileSbar.Max = 2
        End If
        cmdRead.Value = True
    End If
End Sub

Private Sub FileSbar_Change()
    If Not (m_blnReadNextClicked) Then cmdReadFormHere.Value = True
End Sub

Private Sub Form_Load()
    Set m_LogFile = New cLogReader
    Call EnableControls(False)
    Call optEnd_Click
End Sub

Private Sub optEnd_Click()
    m_LogFile.ScanFDirection = False
    ReInitilizeScan
End Sub

Private Sub OptStart_Click()
    m_LogFile.ScanFDirection = True
    ReInitilizeScan
End Sub

Private Sub ReInitilizeScan()
    If (m_LogFile.IsFileOpened) Then
        m_LogFile.fClose
    End If
    txtBuffer.Text = ""
    If (txtFilePath <> "") Then
        cmdReset.Value = True
    End If
End Sub

Private Sub txtFilePath_Change()
    If (Trim(txtFilePath.Text) <> "") Then Call ReadyToRun
End Sub
