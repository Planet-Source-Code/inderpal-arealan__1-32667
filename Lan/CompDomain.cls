VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CompDomain"
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
Private CompDomain As New Collection
Private Declare Function WNetEnumResource Lib "mpr.dll" Alias "WNetEnumResourceA" (ByVal pEnumHwnd As Long, lpcCount As Long, lpBuffer As NETRESOURCE, lpBufferSize As Long) As Long
Private Declare Function WNetCloseEnum Lib "mpr.dll" (ByVal p_lngEnumHwnd As Long) As Long
Private Const NO_ERROR As Long = 0&
Private Const RESOURCETYPE_ANY As Long = &H0&
Private Const RESOURCEUSAGE_ALL As Long = &H0&
Private Const RESOURCE_GLOBALNET As Long = &H2&
Private Const RESOURCE_ENUM_ALL As Long = &HFFFF
Private Const MAX_RESOURCES As Long = 256
Private Declare Function WNetOpenEnum Lib "mpr.dll" Alias "WNetOpenEnumA" (ByVal dwScope As Long, ByVal dwType As Long, ByVal dwUsage As Long, lpNetResource As Any, lppEnumHwnd As Long) As Long
Private Declare Function StrLenA Lib "kernel32" Alias "lstrlenA" (ByVal Ptr As Long) As Long
Private Declare Function StrCopyA Lib "kernel32" Alias "lstrcpyA" (ByVal RetVal As String, ByVal Ptr As Long) As Long

Private Type NETRESOURCE
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    pLocalName As Long
    pRemoteName As Long
    pComment As Long
    pProvider As Long
End Type

Public Function GetCount() As Long
    GetCount = CompDomain.Count
End Function

Public Function GetItem(Index As Long) As String
    GetItem = CompDomain.Item(Index)
End Function

Private Function ComputerDomain(ByVal lPtrToString As Long) As String
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
            ComputerDomain = Left$(sStringValue, lNullPos - 1)
        Else
            ComputerDomain = sStringValue
        End If
    Else
        ComputerDomain = ""
    End If
End Function

Private Sub Class_Initialize()
    Dim tNetAPI(0 To MAX_RESOURCES) As NETRESOURCE
    Dim lBufferSize As Long
    Dim lEnumHwnd As Long
    Dim lReturn As Long
    Dim lCount As Long
    Dim lLoop As Long
    
    Do While CompDomain.Count > 0
        CompDomain.Remove 1
    Loop
    
    lReturn = WNetOpenEnum(RESOURCE_GLOBALNET, RESOURCETYPE_ANY, RESOURCEUSAGE_ALL, tNetAPI(0), lEnumHwnd)
    
    If lReturn = NO_ERROR Then
        lCount = RESOURCE_ENUM_ALL
        lBufferSize = UBound(tNetAPI) * Len(tNetAPI(0))
        lReturn = WNetEnumResource(lEnumHwnd, lCount, tNetAPI(0), lBufferSize)
        
        If lCount > 0 Then
            For lLoop = 0 To lCount - 1
                CompDomain.Add ComputerDomain(tNetAPI(lLoop).pRemoteName)
            Next lLoop
        End If
    End If
    
    If lEnumHwnd <> 0 Then
        Call WNetCloseEnum(lEnumHwnd)
    End If
End Sub

Private Sub Class_Terminate()
    Do While CompDomain.Count > 0
        CompDomain.Remove 1
    Loop
End Sub


