VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CIpHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum OperationalStates
    MIB_IF_OPER_STATUS_NON_OPERATIONAL = 0
    MIB_IF_OPER_STATUS_UNREACHABLE = 1
    MIB_IF_OPER_STATUS_DISCONNECTED = 2
    MIB_IF_OPER_STATUS_CONNECTING = 3
    MIB_IF_OPER_STATUS_CONNECTED = 4
    MIB_IF_OPER_STATUS_OPERATIONAL = 5
End Enum

Public Enum InterfaceTypes
    MIB_IF_TYPE_OTHER = 1
    MIB_IF_TYPE_ETHERNET = 6
    MIB_IF_TYPE_TOKENRING = 9
    MIB_IF_TYPE_FDDI = 15
    MIB_IF_TYPE_PPP = 23
    MIB_IF_TYPE_LOOPBACK = 24
    MIB_IF_TYPE_SLIP = 28
End Enum

Public Enum AdminStatuses
    MIB_IF_ADMIN_STATUS_UP = 1
    MIB_IF_ADMIN_STATUS_DOWN = 2
    MIB_IF_ADMIN_STATUS_TESTING = 3
End Enum

Private Const MAXLEN_IFDESCR = 256
Private Const MAXLEN_PHYSADDR = 8
Private Const MAX_INTERFACE_NAME_LEN = 256

Private Const ERROR_NOT_SUPPORTED = 50&
Private Const ERROR_SUCCESS = 0&


Private Type MIB_IFROW
    wszName(0 To 511) As Byte
    dwIndex As Long             '// index of the interface
    dwType As Long              '// type of interface
    dwMtu As Long               '// max transmission unit
    dwSpeed As Long             '// speed of the interface
    dwPhysAddrLen As Long       '// length of physical address
    bPhysAddr(0 To 7) As Byte   '// physical address of adapter
    dwAdminStatus As Long       '// administrative status
    dwOperStatus As Long        '// operational status
    dwLastChange As Long        '// last time operational status changed
    dwInOctets As Long          '// octets received
    dwInUcastPkts As Long       '// unicast packets received
    dwInNUcastPkts As Long      '// non-unicast packets received
    dwInDiscards As Long        '// received packets discarded
    dwInErrors As Long          '// erroneous packets received
    dwInUnknownProtos As Long   '// unknown protocol packets received
    dwOutOctets As Long         '// octets sent
    dwOutUcastPkts As Long      '// unicast packets sent
    dwOutNUcastPkts As Long     '// non-unicast packets sent
    dwOutDiscards As Long       '// outgoing packets discarded
    dwOutErrors As Long         '// erroneous packets sent
    dwOutQLen As Long           '// output queue length
    dwDescrLen As Long          '// length of bDescr member
    bDescr(0 To 255) As Byte    '// interface description
End Type

Private Declare Function GetIfTable Lib "IPHlpApi" (ByRef pIfRowTable As Any, ByRef pdwSize As Long, ByVal bOrder As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef pDest As Any, ByRef pSource As Any, ByVal Length As Long)

Private mvarInterfaces      As CInterfaces 'local copy

Private m_lngBytesReceived  As Long
Private m_lngBytesSent      As Long

Public Property Set Interfaces(ByVal vData As CInterfaces)
    Set mvarInterfaces = vData
End Property


Public Property Get Interfaces() As CInterfaces
    '
    Set mvarInterfaces = Nothing
    '
    Set mvarInterfaces = New CInterfaces
    Call InitInterfaces(mvarInterfaces)
    '
    Set Interfaces = mvarInterfaces
    '
End Property

Public Property Get BytesReceived() As Long
    BytesReceived = m_lngBytesReceived
End Property

Public Property Get BytesSent() As Long
    BytesSent = m_lngBytesSent
End Property


Private Function InitInterfaces(objInterfaces As CInterfaces) As Boolean
  On Error Resume Next
  '
    Dim arrBuffer()     As Byte
    Dim lngSize         As Long
    Dim lngRetVal       As Long
    Dim lngRows         As Long
    Dim i               As Integer
    Dim j               As Integer
    Dim IfRowTable      As MIB_IFROW
    Dim objInterface    As New CInterface
    '
    lngSize = 0
    '
    'Reset the BytesReceived and BytesSent properties
    '
    m_lngBytesReceived = 0
    m_lngBytesSent = 0
    '
    'Call the GetIfTable just to get the buffer size into the lngSize variable
    lngRetVal = GetIfTable(ByVal 0&, lngSize, 0)
    '
    If lngRetVal = ERROR_NOT_SUPPORTED Then
        '
        'This API works only on Win 98/2000 and NT4 with SP4
        MsgBox "IP Helper is not supported by this system."
        Exit Function
        '
    End If
    '
    'Prepare the buffer
    ReDim arrBuffer(0 To lngSize - 1) As Byte
    '
    'And call the function one more time
    lngRetVal = GetIfTable(arrBuffer(0), lngSize, 0)
    '
    If lngRetVal = ERROR_SUCCESS Then
        '
        'The first 4 bytes (the Long value) contain the quantity of the table rows
        'Get that value into the lngRows variable
        CopyMemory lngRows, arrBuffer(0), 4
        '
        For i = 1 To lngRows
            '
            'Copy the table row data to the IfRowTable structure
            CopyMemory IfRowTable, arrBuffer(4 + (i - 1) * Len(IfRowTable)), Len(IfRowTable)
            '
            With IfRowTable
                '
                objInterface.InterfaceDescription = Left(StrConv(.bDescr, vbUnicode), .dwDescrLen)
                '
                If .dwPhysAddrLen > 0 Then
                    For j = 0 To .dwPhysAddrLen - 1
                        objInterface.AdapterAddress = objInterface.AdapterAddress & _
                                                  CStr(IIf(.bPhysAddr(j) = 0, "00", Hex(.bPhysAddr(j)))) & "-"
                        '
                    Next j
                objInterface.AdapterAddress = Left(objInterface.AdapterAddress, Len(objInterface.AdapterAddress) - 1)
                    
                End If
                '
                objInterface.AdminStatus = .dwAdminStatus
                objInterface.InterfaceIndex = .dwIndex
                objInterface.DiscardedIncomingPackets = .dwInDiscards
                objInterface.IncomingErrors = .dwInErrors
                objInterface.NonunicastPacketsReceived = .dwInNUcastPkts
                objInterface.OctetsReceived = .dwInOctets
                objInterface.UnicastPacketsReceived = .dwInUcastPkts
                objInterface.UnknownProtocolPackets = .dwInUnknownProtos
                objInterface.LastChange = .dwLastChange
                objInterface.MaximumTransmissionUnit = .dwMtu
                objInterface.OperationalStatus = .dwOperStatus
                objInterface.DiscardedOutgoingPackets = .dwOutDiscards
                objInterface.OutgoingErrors = .dwOutErrors
                objInterface.NonunicastPacketsSent = .dwOutNUcastPkts
                objInterface.OctetsSent = .dwOutOctets
                objInterface.OutputQueueLength = .dwOutQLen
                objInterface.UnicastPacketsSent = .dwOutUcastPkts
                objInterface.Speed = .dwSpeed
                objInterface.InterfaceType = .dwType
                objInterface.InterfaceName = StrConv(.wszName, vbUnicode)
                '
                'Collect traffic info for all the interfaces
                '
                m_lngBytesReceived = m_lngBytesReceived + .dwInOctets
                m_lngBytesSent = m_lngBytesSent + .dwOutOctets
                '
            End With
            '
            mvarInterfaces.Add objInterface
            '
        Next i
        '
    End If
    '
End Function

