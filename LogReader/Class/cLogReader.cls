VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLogReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'***********************************************************************************
'FILENAME       :   cFile.cls

'AUTHOR         :   Rajneesh Noonia

'VERSION HISTORY:
'                   Date        Who             Version     Description
'                   12/04/05    Rajneesh Noonia    1.0      Implementation

'DESCRIPTION    :   This class implements functions to handle text files without
'                   using file system object or any other references.
'
'***********************************************************************************

Private m_FileHandle As Long

Private m_lngPosition As Long

Private m_lngBufferSize As Long
Private m_blnScanForward As Boolean
Private m_lngMaxPointer As Long

Public Property Get IsFileOpened() As Boolean
    IsFileOpened = m_FileHandle
End Property
'Open File for Read / Write
Public Function OpenLogFile(ByVal FilePath As String) As Boolean
    On Error GoTo ErrorTrap
    
    Dim pdblFraction As Double
    'Close if Already open
    If (m_FileHandle <> 0) Then Me.fClose
    
    m_lngBufferSize = 1000
    
    
    'm_blnScanForward = True
    'Get Next File Handle
    m_FileHandle = FreeFile()
    Open FilePath For Random As #m_FileHandle Len = m_lngBufferSize
    pdblFraction = FileSystem.LOF(m_FileHandle) / m_lngBufferSize
    
    If (pdblFraction > Int(pdblFraction)) Then
        m_lngMaxPointer = Int(pdblFraction) + 1
    Else
        m_lngMaxPointer = Int(pdblFraction)
    End If
    If (m_blnScanForward) Then
        m_lngPosition = 1
    Else
        m_lngPosition = m_lngMaxPointer
    End If
    OpenLogFile = True
    Exit Function
ErrorTrap:
    OpenLogFile = False
End Function


Public Function FileExist(ByVal Filename As String) As Boolean
    On Error GoTo ErrorTrap
    Dim lngTemp As Long
    
    FileExist = False
    
    lngTemp = FreeFile
    Open Filename For Input As #lngTemp
    Close #lngTemp
    
    FileExist = True
    Exit Function
ErrorTrap:

    If (Err.Number = 55) Then 'File ALREADY OPEN
        FileExist = True
        Err.Clear
        Exit Function
    Else
        FileExist = False
        Err.Clear
        Exit Function
    End If
    
End Function

'Read Line From File
Public Function ReadBuffer() As String
On Error GoTo ErrorTrap
   Dim pBuffer As String * 1000
   Dim pptr As Long
   
    Get #m_FileHandle, m_lngPosition, pBuffer
    
    If (m_blnScanForward) Then
        m_lngPosition = m_lngPosition + 1
    Else
        m_lngPosition = m_lngPosition - 1
    End If
    ReadBuffer = pBuffer
   Exit Function
ErrorTrap:
    Err.Raise Err.Number, "cFile.ReadLine", Err.Description
End Function

Public Property Let ScanFDirection(ByVal ScanFarward As Boolean)
    m_blnScanForward = ScanFarward
    If (m_blnScanForward) Then
        m_lngPosition = 1
    Else
        m_lngPosition = m_lngMaxPointer
    End If
End Property

Public Property Get ScanFDirection() As Boolean
    ScanFDirection = m_blnScanForward
End Property

Public Property Let PointerPosition(ByVal Position As Long)
    m_lngPosition = Position
End Property

Public Property Get PointerPosition() As Long
    PointerPosition = m_lngPosition
End Property

Public Property Let BufferSize(ByVal Size As Long)
    m_lngBufferSize = Size
End Property

Public Property Get BufferSize() As Long
    BufferSize = m_lngBufferSize
End Property

Public Sub MoveToEnd()
    m_lngPosition = m_lngMaxPointer
End Sub

Public Property Get MaxPosition() As Long
    MaxPosition = m_lngMaxPointer
End Property


'Get Pointer Status
Public Property Get EOF() As Boolean
    EOF = True
    If (m_FileHandle <> 0) Then
        If (m_lngMaxPointer >= m_lngPosition And m_lngPosition >= 1) Then EOF = False
    End If
End Property

Public Property Get BOF() As Boolean
    BOF = m_lngPosition > 0
End Property

'Close the File
Public Sub fClose()
    If (m_FileHandle <> 0) Then
        Close #m_FileHandle
        m_FileHandle = 0
    End If
End Sub

Private Sub Class_Initialize()
    '
    m_FileHandle = 0
End Sub

Private Sub Class_Terminate()
    'Realese Resources
    If (m_FileHandle <> 0) Then
        Call Me.fClose
    End If
End Sub
