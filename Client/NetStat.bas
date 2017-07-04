Attribute VB_Name = "NetStat"
Option Explicit

'================= TCP things ====================
'state of the connection
Public Const MIB_TCP_STATE_CLOSED = 0
Public Const MIB_TCP_STATE_LISTEN = 1
Public Const MIB_TCP_STATE_SYN_SENT = 2
Public Const MIB_TCP_STATE_SYN_RCVD = 3
Public Const MIB_TCP_STATE_ESTAB = 4
Public Const MIB_TCP_STATE_FIN_WAIT1 = 5
Public Const MIB_TCP_STATE_FIN_WAIT2 = 6
Public Const MIB_TCP_STATE_CLOSE_WAIT = 7
Public Const MIB_TCP_STATE_CLOSING = 8
Public Const MIB_TCP_STATE_LAST_ACK = 9
Public Const MIB_TCP_STATE_TIME_WAIT = 10
Public Const MIB_TCP_STATE_DELETE_TCB = 11

Type MIB_TCPROW
  dwState As Long        'state of the connection
  dwLocalAddr As String * 4    'address on local computer
  dwLocalPort As String * 4    'port number on local computer
  dwRemoteAddr As String * 4   'address on remote computer
  dwRemotePort As String * 4   'port number on remote computer
End Type

Type MIB_TCPTABLE
  dwNumEntries As Long    'number of entries in the table
  table(100) As MIB_TCPROW   'array of TCP connections
End Type

Declare Function GetTcpTable Lib "IPhlpAPI" _
  (pTcpTable As MIB_TCPTABLE, pdwSize As Long, bOrder As Long) As Long

'================= UDP things ====================
Type MIB_UDPROW
  dwLocalAddr As String * 4 'address on local computer
  dwLocalPort As String * 4 'port number on local computer
End Type

Type MIB_UDPTABLE
  dwNumEntries As Long    'number of entries in the table
  table(100) As MIB_UDPROW   'table of MIB_UDPROW structs
End Type

Declare Function GetUdpTable Lib "IPhlpAPI" _
  (pUdpTable As MIB_UDPTABLE, pdwSize As Long, bOrder As Long) As Long

'================= Statistics ====================
Type MIB_IPSTATS
  dwForwarding As Long       ' IP forwarding enabled or disabled
  dwDefaultTTL As Long       ' default time-to-live
  dwInReceives As Long       ' datagrams received
  dwInHdrErrors As Long      ' received header errors
  dwInAddrErrors As Long     ' received address errors
  dwForwDatagrams As Long    ' datagrams forwarded
  dwInUnknownProtos As Long  ' datagrams with unknown protocol
  dwInDiscards As Long       ' received datagrams discarded
  dwInDelivers As Long       ' received datagrams delivered
  dwOutRequests As Long      '
  dwRoutingDiscards As Long  '
  dwOutDiscards As Long      ' sent datagrams discarded
  dwOutNoRoutes As Long      ' datagrams for which no route
  dwReasmTimeout As Long     ' datagrams for which all frags didn't arrive
  dwReasmReqds As Long       ' datagrams requiring reassembly
  dwReasmOks As Long         ' successful reassemblies
  dwReasmFails As Long       ' failed reassemblies
  dwFragOks As Long          ' successful fragmentations
  dwFragFails As Long        ' failed fragmentations
  dwFragCreates As Long      ' datagrams fragmented
  dwNumIf As Long           ' number of interfaces on computer
  dwNumAddr As Long         ' number of IP address on computer
  dwNumRoutes As Long       ' number of routes in routing table
End Type

Declare Function GetIpStatistics Lib "IPhlpAPI" _
  (pStats As MIB_IPSTATS) As Long

Type MIBICMPSTATS
  dwMsgs As Long            ' number of messages
  dwErrors As Long          ' number of errors
  dwDestUnreachs As Long    ' destination unreachable messages
  dwTimeExcds As Long       ' time-to-live exceeded messages
  dwParmProbs As Long       ' parameter problem messages
  dwSrcQuenchs As Long      ' source quench messages
  dwRedirects As Long       ' redirection messages
  dwEchos As Long           ' echo requests
  dwEchoReps As Long        ' echo replies
  dwTimestamps As Long      ' timestamp requests
  dwTimestampReps As Long   ' timestamp replies
  dwAddrMasks As Long       ' address mask requests
  dwAddrMaskReps As Long    ' address mask replies
End Type

Type MIBICMPINFO
  icmpInStats As MIBICMPSTATS        ' stats for incoming messages
  icmpOutStats As MIBICMPSTATS       ' stats for outgoing messages
End Type

Declare Function GetIcmpStatistics Lib "IPhlpAPI" _
  (pStats As MIBICMPINFO) As Long

Type MIB_TCPSTATS
  dwRtoAlgorithm As Long    ' timeout algorithm
  dwRtoMin As Long          ' minimum timeout
  dwRtoMax As Long          ' maximum timeout
  dwMaxConn As Long         ' maximum connections
  dwActiveOpens As Long     ' active opens
  dwPassiveOpens As Long    ' passive opens
  dwAttemptFails As Long    ' failed attempts
  dwEstabResets As Long     ' establised connections reset
  dwCurrEstab As Long       ' established connections
  dwInSegs As Long          ' segments received
  dwOutSegs As Long         ' segment sent
  dwRetransSegs As Long     ' segments retransmitted
  dwInErrs As Long          ' incoming errors
  dwOutRsts As Long         ' outgoing resets
  dwNumConns As Long        ' cumulative connections
End Type

Declare Function GetTcpStatistics Lib "IPhlpAPI" _
  (pStats As MIB_TCPSTATS) As Long

Type MIB_UDPSTATS
  dwInDatagrams As Long    ' received datagrams
  dwNoPorts As Long        ' datagrams for which no port
  dwInErrors As Long       ' errors on received datagrams
  dwOutDatagrams As Long   ' sent datagrams
  dwNumAddrs As Long       ' number of entries in UDP listener table
End Type

Declare Function GetUdpStatistics Lib "IPhlpAPI" _
  (pStats As MIB_UDPSTATS) As Long

'================= Conversion ====================
Function c_port(s) As Long
On Error Resume Next
  c_port = Asc(Mid(s, 1, 1)) * 256 + Asc(Mid(s, 2, 1))
End Function

Function c_ip(s) As String
  c_ip = Asc(Mid(s, 1, 1)) & "." & Asc(Mid(s, 2, 1)) & "." & Asc(Mid(s, 3, 1)) & "." & Asc(Mid(s, 4, 1))
End Function

Function c_state(s) As String
  Select Case s
  Case MIB_TCP_STATE_CLOSED: c_state = "CLOSED"
  Case MIB_TCP_STATE_LISTEN: c_state = "LISTEN"
  Case MIB_TCP_STATE_SYN_SENT: c_state = "SYN_SENT"
  Case MIB_TCP_STATE_SYN_RCVD: c_state = "SYN_RCVD"
  Case MIB_TCP_STATE_ESTAB: c_state = "ESTAB"
  Case MIB_TCP_STATE_FIN_WAIT1: c_state = "FIN_WAIT1"
  Case MIB_TCP_STATE_FIN_WAIT2: c_state = "FIN_WAIT2"
  Case MIB_TCP_STATE_CLOSE_WAIT: c_state = "CLOSE_WAIT"
  Case MIB_TCP_STATE_CLOSING: c_state = "CLOSING"
  Case MIB_TCP_STATE_LAST_ACK: c_state = "LAST_ACK"
  Case MIB_TCP_STATE_TIME_WAIT: c_state = "TIME_WAIT"
  Case MIB_TCP_STATE_DELETE_TCB: c_state = "DELETE_TCB"
  Case Else: c_state = "UNDEFINED"
  End Select
End Function

