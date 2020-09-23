VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CompName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
     '########################################'
     '   Programmed By Inderpal Singh         '
     '   Email: inderpal0@hotmail.com         '
     '   Date: March 03, 2002                 '
     '   Homepage: http://connect.to/lanserver'
     '########################################'
Option Explicit

Private CompName As New Collection
Private sDomain As String
Private Const GMEM_FIXED = &H0
Private Const RESOURCE_GLOBALNET = &H2
Private Const RESOURCETYPE_DISK = &H1
Private Const GMEM_ZEROINIT = &H40
Private Const ERROR_MORE_DATA = 234
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function WNetOpenEnum Lib "mpr.dll" Alias "WNetOpenEnumA" (ByVal dwScope As Long, ByVal dwType As Long, ByVal dwUsage As Long, lpNetResource As Any, lphEnum As Long) As Long

Private Declare Function WNetCloseEnum Lib "mpr.dll" (ByVal hEnum As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function WNetEnumResource Lib "mpr.dll" Alias "WNetEnumResourceA" (ByVal hEnum As Long, lpcCount As Long, ByVal lpBuffer As Long, ByRef lpBufferSize As Long) As Long
Private Declare Function StrLenA Lib "kernel32" Alias "lstrlenA" (ByVal Ptr As Long) As Long
Private Declare Function StrCopyA Lib "kernel32" Alias "lstrcpyA" (ByVal RetVal As String, ByVal Ptr As Long) As Long
Private Type NETRES2
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    lpLocalName As String
    lpRemoteName As String
    lpComment As String
    lpProvider As String
End Type
Private Type NETRESOURCE    'Structure
    dwScope As Long         'for information
    dwType As Long          'about
    dwDisplayType As Long   'Network
    dwUsage As Long         'Resource
    lpLocalName As Long
    lpRemoteName As Long
    lpComment As Long
    lpProvider As Long
End Type
Public Property Get Domain() As String
    Domain = sDomain
End Property

Public Property Let Domain(Value As String)
    sDomain = Value
End Property

Public Function GetCount() As Long
    GetCount = CompName.Count
End Function

Public Function GetItem(Index As Long) As String
    GetItem = CompName.Item(Index)
End Function
Public Sub Refresh()
    Dim lBufferPtrTemp As Long
    Dim CName As String
    Dim tNetRes As NETRES2
    Dim tNR As NETRESOURCE
    Dim lBufferPtr As Long
    Dim lEnumHwnd As Long
    Dim lReturn As Long
    Dim lBuffer As Long
    Dim lCount As Long
    Dim lK As Long
    
    With tNetRes
        .lpRemoteName = sDomain
        .dwDisplayType = 1
    End With
    
    lReturn = WNetOpenEnum(RESOURCE_GLOBALNET, RESOURCETYPE_DISK, 0, tNetRes, lEnumHwnd)
           
    lBuffer = 1024 * 10
    lBufferPtr = GlobalAlloc(GPTR, lBuffer)
    Do
        lCount = -1
        lReturn = WNetEnumResource(lEnumHwnd, lCount, lBufferPtr, lBuffer)
            If lReturn = 0 Then
                lBufferPtrTemp = lBufferPtr
                For lK = 1 To lCount
                    CopyMemory tNR, ByVal lBufferPtrTemp, LenB(tNR)
                    CName = ComputerName(tNR.lpRemoteName)
                    If CName <> "" Then
                        CName = Mid(CName, InStr(CName, "\\") + 2)
                        CompName.Add CName
                    End If
                    lBufferPtrTemp = lBufferPtrTemp + LenB(tNR)
                Next
            End If
    Loop Until lCount = 0
    If lEnumHwnd <> 0 Then
        lReturn = WNetCloseEnum(lEnumHwnd)
    End If
    GlobalFree lBufferPtr
End Sub

Private Function ComputerName(ByVal lPtrToString As Long) As String
    On Local Error Resume Next
    Dim lLength As Long
    Dim sStringValue As String
    Dim lNullPos As Long
    Dim lReturn As Long
    
    lLength = StrLenA(lPtrToString)
    If lPtrToString > 0 And lLength > 0 Then
        sStringValue = Space$(lLength + 1)
        lReturn = StrCopyA(sStringValue, lPtrToString)
        lNullPos = InStr(1, sStringValue, Chr$(0))
        If lNullPos > 0 Then
            ComputerName = Left$(sStringValue, lNullPos - 1)
        Else
            ComputerName = sStringValue
        End If
    Else
        ComputerName = ""
    End If
End Function

