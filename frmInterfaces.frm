VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInterfaces 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Network Interfaces Information"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6480
   Icon            =   "frmInterfaces.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   6480
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   5160
      TabIndex        =   10
      Top             =   720
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "Traffic Summary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   6240
      Width           =   4935
      Begin VB.Label lblSent 
         Caption         =   "000 000 000"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   9
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblRecv 
         Caption         =   "000 000 000"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Bytes sent:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1800
         TabIndex        =   7
         Top             =   480
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Bytes received:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1800
         TabIndex        =   6
         Top             =   240
         Width           =   1125
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "frmInterfaces.frx":030A
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   5160
      Top             =   1440
   End
   Begin VB.Frame Frame2 
      Caption         =   "Interface Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   4935
      Begin MSComctlLib.ListView lvInterfaceInfo 
         Height          =   4935
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   8705
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Select Interface"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4935
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmInterfaces.frx":0614
         Left            =   120
         List            =   "frmInterfaces.frx":0616
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "&Hide"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   5520
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInterfaces.frx":0618
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInterfaces.frx":0934
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInterfaces.frx":0C50
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInterfaces.frx":0F6C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5520
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInterfaces.frx":1288
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInterfaces.frx":13E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInterfaces.frx":1540
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInterfaces.frx":169C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mPopupSys 
      Caption         =   "&SysTray"
      Visible         =   0   'False
      Begin VB.Menu mPopRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mPopExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmInterfaces"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'user defined type required by Shell_NotifyIcon API call
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

'constants required by Shell_NotifyIcon API call:
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201     'Button down
Private Const WM_LBUTTONUP = &H202       'Button up
Private Const WM_LBUTTONDBLCLK = &H203   'Double-click
Private Const WM_RBUTTONDOWN = &H204     'Button down
Private Const WM_RBUTTONUP = &H205       'Button up
Private Const WM_RBUTTONDBLCLK = &H206   'Double-click

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Private Const MAXLEN_IFDESCR = 256
Private Const MAXLEN_PHYSADDR = 8
Private Const MAX_INTERFACE_NAME_LEN = 256

Private nid As NOTIFYICONDATA
Private m_objIpHelper As CIpHelper

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdHide_Click()
    Me.WindowState = vbMinimized
End Sub

Private Sub Combo1_Click()
    Call UpdateInterfaceInfo(Combo1.ListIndex + 1)
End Sub

Private Sub Form_Load()
    '
    Dim objInterface As CInterface
    '
    'Configure the listview control
    '
    'Add column headers
    '
    lvInterfaceInfo.ColumnHeaders.Add , , "Parameter", 3000
    lvInterfaceInfo.ColumnHeaders.Add , , "Value", 1600
    '
    'Add listview items - interface parameters
    '
    With lvInterfaceInfo.ListItems
        .Add , , "Name of the interface"
        .Add , , "Index of the interface"
        .Add , , "Type of interface"
        .Add , , "Max transmission unit"
        .Add , , "Speed of the interface"
        .Add , , "Physical address of adapter"
        .Add , , "Administrative status"
        .Add , , "Operational status"
        .Add , , "Last time operational status changed"
        .Add , , "Octets received"
        .Add , , "Unicast packets received"
        .Add , , "Non-unicast packets received"
        .Add , , "Received packets discarded"
        .Add , , "Erroneous packets received"
        .Add , , "Unknown protocol packets received"
        .Add , , "Octets sent"
        .Add , , "Unicast packets sent"
        .Add , , "Non-unicast packets sent"
        .Add , , "Outgoing packets discarded"
        .Add , , "Erroneous packets sent"
        .Add , , "Output queue length"
    End With
    '
    'Create an instance of the CIpHelper class
    Set m_objIpHelper = New CIpHelper
    '
    'Add descriptions of the network interfaces into the listbox control
    For Each objInterface In m_objIpHelper.Interfaces
        Combo1.AddItem objInterface.InterfaceDescription
    Next
    '
    'Define selected item in the listbox control
    Combo1.ListIndex = 0
    '
    'The system tray code
    '
    'the form must be fully visible before calling Shell_NotifyIcon
    Me.Show
    Me.Refresh
    With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = ImageList1.ListImages(4).Picture
        .szTip = "Bytes received: " & m_objIpHelper.BytesReceived & _
                 " Bytes sent: " & m_objIpHelper.BytesSent & vbNullChar
    End With
    '
    Shell_NotifyIcon NIM_ADD, nid
    '
End Sub

Private Sub UpdateInterfaceInfo(intIndex)
    '
    Dim objInterface        As CInterface
    Static st_objInterface  As CInterface
    '
    Static lngBytesRecv     As Long
    Static lngBytesSent     As Long
    '
    Dim blnIsRecv           As Boolean
    Dim blnIsSent           As Boolean
    '
    If intIndex > m_objIpHelper.Interfaces.Count Then
        Exit Sub
    End If
    '
    If st_objInterface Is Nothing Then Set st_objInterface = New CInterface
    '
    Set objInterface = m_objIpHelper.Interfaces(intIndex)
    '
    With lvInterfaceInfo.ListItems
        If Not objInterface.InterfaceName = st_objInterface.InterfaceName Then _
            .Item(1).SubItems(1) = objInterface.InterfaceName
        If Not objInterface.InterfaceIndex = st_objInterface.InterfaceIndex Then _
            .Item(2).SubItems(1) = objInterface.InterfaceIndex
        If Not objInterface.InterfaceType = st_objInterface.InterfaceType Then
            Select Case objInterface.InterfaceType
                Case MIB_IF_TYPE_ETHERNET: .Item(3).SubItems(1) = "Ethernet"
                Case MIB_IF_TYPE_FDDI: .Item(3).SubItems(1) = "FDDI"
                Case MIB_IF_TYPE_LOOPBACK: .Item(3).SubItems(1) = "Loopback"
                Case MIB_IF_TYPE_OTHER: .Item(3).SubItems(1) = "Other"
                Case MIB_IF_TYPE_PPP: .Item(3).SubItems(1) = "PPP"
                Case MIB_IF_TYPE_SLIP: .Item(3).SubItems(1) = "SLIP"
                Case MIB_IF_TYPE_TOKENRING: .Item(3).SubItems(1) = "TokenRing"
            End Select
        End If
        If Not objInterface.MaximumTransmissionUnit = st_objInterface.MaximumTransmissionUnit Then _
            .Item(4).SubItems(1) = objInterface.MaximumTransmissionUnit
        If Not objInterface.Speed = st_objInterface.Speed Then _
            .Item(5).SubItems(1) = Trim(Format(objInterface.Speed, "### ### ###"))
        If Not objInterface.AdapterAddress = st_objInterface.AdapterAddress Then _
            .Item(6).SubItems(1) = objInterface.AdapterAddress
        If Not objInterface.AdminStatus = st_objInterface.AdminStatus Then
            Select Case objInterface.AdminStatus
                Case MIB_IF_ADMIN_STATUS_DOWN: .Item(7).SubItems(1) = "Down"
                Case MIB_IF_ADMIN_STATUS_TESTING: .Item(7).SubItems(1) = "Testing"
                Case MIB_IF_ADMIN_STATUS_UP: .Item(7).SubItems(1) = "Up"
            End Select
        End If
        If Not objInterface.OperationalStatus = st_objInterface.OperationalStatus Then
            Select Case objInterface.OperationalStatus
                Case MIB_IF_OPER_STATUS_CONNECTED: .Item(8).SubItems(1) = "Connected"
                Case MIB_IF_OPER_STATUS_CONNECTING: .Item(8).SubItems(1) = "Connecting"
                Case MIB_IF_OPER_STATUS_DISCONNECTED: .Item(8).SubItems(1) = "Disconnected"
                Case MIB_IF_OPER_STATUS_NON_OPERATIONAL: .Item(8).SubItems(1) = "Non operational"
                Case MIB_IF_OPER_STATUS_OPERATIONAL: .Item(8).SubItems(1) = "Operational"
                Case MIB_IF_OPER_STATUS_UNREACHABLE: .Item(8).SubItems(1) = "Unreachable"
            End Select
        End If
        If Not objInterface.LastChange = st_objInterface.LastChange Then _
            .Item(9).SubItems(1) = objInterface.LastChange
        If Not objInterface.OctetsReceived = st_objInterface.OctetsReceived Then _
            .Item(10).SubItems(1) = Trim(Format(objInterface.OctetsReceived, "### ### ### ###"))
        If Not objInterface.UnicastPacketsReceived = st_objInterface.UnicastPacketsReceived Then _
            .Item(11).SubItems(1) = objInterface.UnicastPacketsReceived
        If Not objInterface.NonunicastPacketsReceived = st_objInterface.NonunicastPacketsReceived Then _
            .Item(12).SubItems(1) = objInterface.NonunicastPacketsReceived
        If Not objInterface.DiscardedIncomingPackets = st_objInterface.DiscardedIncomingPackets Then _
            .Item(13).SubItems(1) = objInterface.DiscardedIncomingPackets
        If Not objInterface.IncomingErrors = st_objInterface.IncomingErrors Then _
            .Item(14).SubItems(1) = objInterface.IncomingErrors
        If Not objInterface.UnknownProtocolPackets = st_objInterface.UnknownProtocolPackets Then _
            .Item(15).SubItems(1) = objInterface.UnknownProtocolPackets
        If Not objInterface.OctetsSent = st_objInterface.OctetsSent Then _
            .Item(16).SubItems(1) = Trim(Format(objInterface.OctetsSent, "### ### ### ###"))
        If Not objInterface.UnicastPacketsSent = st_objInterface.UnicastPacketsSent Then _
            .Item(17).SubItems(1) = objInterface.UnicastPacketsSent
        If Not objInterface.NonunicastPacketsSent = st_objInterface.NonunicastPacketsSent Then _
            .Item(18).SubItems(1) = objInterface.NonunicastPacketsSent
        If Not objInterface.DiscardedOutgoingPackets = st_objInterface.DiscardedOutgoingPackets Then _
            .Item(19).SubItems(1) = objInterface.DiscardedOutgoingPackets
        If Not objInterface.OutgoingErrors = st_objInterface.OutgoingErrors Then _
            .Item(20).SubItems(1) = objInterface.OutgoingErrors
        If Not objInterface.OutputQueueLength = st_objInterface.OutputQueueLength Then _
            .Item(21).SubItems(1) = objInterface.OutputQueueLength
    End With
    '
    lblRecv.Caption = Trim(Format(m_objIpHelper.BytesReceived, "### ### ### ###"))
    lblSent.Caption = Trim(Format(m_objIpHelper.BytesSent, "### ### ### ###"))
    '
    blnIsRecv = (m_objIpHelper.BytesReceived > lngBytesRecv)
    blnIsSent = (m_objIpHelper.BytesSent > lngBytesSent)
    '
    If blnIsRecv And blnIsSent Then
        Set Image1.Picture = ImageList2.ListImages(4).Picture
        nid.hIcon = ImageList1.ListImages(4).Picture
    ElseIf (Not blnIsRecv) And blnIsSent Then
        Set Image1.Picture = ImageList2.ListImages(3).Picture
        nid.hIcon = ImageList1.ListImages(3).Picture
    ElseIf blnIsRecv And (Not blnIsSent) Then
        Set Image1.Picture = ImageList2.ListImages(2).Picture
        nid.hIcon = ImageList1.ListImages(2).Picture
    ElseIf Not (blnIsRecv And blnIsSent) Then
        Set Image1.Picture = ImageList2.ListImages(1).Picture
        nid.hIcon = ImageList1.ListImages(1).Picture
    End If
    '
    lngBytesRecv = m_objIpHelper.BytesReceived
    lngBytesSent = m_objIpHelper.BytesSent
    '
    nid.szTip = "Bytes received: " & lngBytesRecv & " Bytes sent: " & lngBytesSent & vbNullChar
    Shell_NotifyIcon NIM_MODIFY, nid
    '
    Set st_objInterface = objInterface
    '
End Sub

Private Sub Timer1_Timer()
    Call UpdateInterfaceInfo(Combo1.ListIndex + 1)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '
    'this procedure receives the callbacks from the System Tray icon.
    '
    Dim Result As Long
    Dim msg As Long
    '
    'the value of X will vary depending upon the scalemode setting
    '
    If Me.ScaleMode = vbPixels Then
        msg = X
    Else
        msg = X / Screen.TwipsPerPixelX
    End If
    '
    Select Case msg
        Case WM_LBUTTONUP        '514 restore form window
            Me.WindowState = vbNormal
            Result = SetForegroundWindow(Me.hwnd)
            Me.Show
        Case WM_LBUTTONDBLCLK    '515 restore form window
            Me.WindowState = vbNormal
            Result = SetForegroundWindow(Me.hwnd)
            Me.Show
        Case WM_RBUTTONUP        '517 display popup menu
            Result = SetForegroundWindow(Me.hwnd)
            Me.PopupMenu Me.mPopupSys
    End Select
    
End Sub

Private Sub Form_Resize()
    '
    'this is necessary to assure that the minimized window is hidden
    '
    If Me.WindowState = vbMinimized Then Me.Hide
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '
    'this removes the icon from the system tray
    '
    Shell_NotifyIcon NIM_DELETE, nid
    '
End Sub

Private Sub mPopExit_Click()
    '
    'called when user clicks the popup menu Exit command
    '
    Unload Me
    '
End Sub


Private Sub mPopRestore_Click()
    '
    'called when the user clicks the popup menu Restore command
    '
    Me.WindowState = vbNormal
    Call SetForegroundWindow(Me.hwnd)
    Me.Show
    '
End Sub
