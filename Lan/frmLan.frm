VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmLan 
   ClientHeight    =   2745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmLan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TimerScroll 
      Interval        =   200
      Left            =   3480
      Top             =   1560
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   2280
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid LanGrid 
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   3625
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
   End
End
Attribute VB_Name = "frmLan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
     '########################################'
     '   Programmed By Inderpal Singh         '
     '   Email: inderpal0@hotmail.com         '
     '   Date: March 03, 2002                 '
     '   Homepage: http://connect.to/lanserver'
     '########################################'

Const Domain = 0
Const Name1 = 1
Const IpAddress = 2
Dim strIpAddress    As String
Dim Name2 As String
Private Sub Form_Load()
    Dim lngRetVal      As Long
    Dim strErrorMsg    As String
    Dim udtWinsockData As WSAData
    lngRetVal = WSAStartup(&H101, udtWinsockData)
    cmdRefresh_Click
End Sub

Private Sub cmdRefresh_Click()
    Dim CompName As New CompName
    Dim CompDomain As New CompDomain
    Dim lCurrentNode As Long
    Dim Add As Integer
    Dim ip As Long
    Dim ips As Long
    
    LanGrid.Clear
    LanGrid.Rows = 2
    LanGrid.ColAlignment(Domain) = flexAlignLeftCenter
    LanGrid.ColAlignment(Name1) = flexAlignCenterCenter
    LanGrid.ColAlignment(IpAddress) = flexAlignCenterCenter
    LanGrid.ColWidth(0) = 1000
    LanGrid.ColWidth(1) = 2040
    LanGrid.ColWidth(2) = 1300
    LanGrid.Row = 0
    LanGrid.Col = 0
    LanGrid.Text = "Workgroup"
    LanGrid.Col = 1
    LanGrid.Text = "Computer Name"
    LanGrid.Col = 2
    LanGrid.Text = "Ip Address"
    LanGrid.SelectionMode = flexSelectionByRow
    For ip = 1 To CompDomain.GetCount
        CompName.Domain = CompDomain.GetItem(ip)
        CompName.Refresh
        For ips = 1 To CompName.GetCount
            LanGrid.Row = 1
            If LanGrid.Text = "" Then
                LanGrid.Col = Domain
                LanGrid.Text = CompDomain.GetItem(ip)
                LanGrid.Col = Name1
                Name2 = CompName.GetItem(ips)
                LanGrid.Text = Name2
                Call CompIp(Name2)
                LanGrid.Col = IpAddress
                LanGrid.Text = strIpAddress
                strIpAddress = ""
            Else
                LanGrid.AddItem CompDomain.GetItem(ip)
                Add = LanGrid.Rows - 1
                LanGrid.Row = Add
                LanGrid.Col = Name1
                Name2 = CompName.GetItem(ips)
                LanGrid.Text = Name2
                Call CompIp(Name2)
                LanGrid.Col = IpAddress
                LanGrid.Text = strIpAddress
                strIpAddress = ""
            End If
        Next ips
    Next ip
End Sub

Private Sub CompIp(Name As String)
    Dim lngPtrToHOSTENT As Long
    Dim udtHostent      As HOSTENT
    Dim lngPtrToIP      As Long
    Dim arrIpAddress()  As Byte
    Dim Name3 As Long
    lngPtrToHOSTENT = gethostbyname(Trim$(Name))
    If lngPtrToHOSTENT = 0 Then
        strIpAddress = "Not Open"
    Else
        RtlMoveMemory udtHostent, lngPtrToHOSTENT, LenB(udtHostent)
        RtlMoveMemory lngPtrToIP, udtHostent.hAddrList, 4
        Do Until lngPtrToIP = 0
            ReDim arrIpAddress(1 To udtHostent.hLength)
            RtlMoveMemory arrIpAddress(1), lngPtrToIP, udtHostent.hLength
            For i = 1 To udtHostent.hLength
                strIpAddress = strIpAddress & arrIpAddress(i) & "."
            Next
            strIpAddress = Left$(strIpAddress, Len(strIpAddress) - 1)
            udtHostent.hAddrList = udtHostent.hAddrList + LenB(udtHostent.hAddrList)
            RtlMoveMemory lngPtrToIP, udtHostent.hAddrList, 4
         Loop
    End If
End Sub

Private Sub TimerScroll_Timer()
    Static a As Integer
    a = a + 1
    Me.Caption = Mid("Local Area Network", 1, a)
    If a = 18 Then a = 0
End Sub
