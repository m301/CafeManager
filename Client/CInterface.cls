VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private m_strInterfaceName As String
Private m_lngInterfaceIndex As Long
Private m_InterfaceType As InterfaceTypes
Private m_lngSpeed As Long
Private m_lngAdapterAddress As String
Private m_AdminStatus As AdminStatuses
Private m_OperationalStatus As OperationalStates
Private m_datLastChange As Date
Private m_lngOctetsReceived As Long
Private m_lngUnicastPacketsReceived As Long
Private m_lngMaximumTransmissionUnit As Long
Private m_lngNonunicastPacketsReceived As Long
Private m_lngDiscardedIncomingPackets As Long
Private m_lngIncomingErrors As Long
Private m_lngUnknownProtocolPackets As Long
Private m_lngOctetsSent As Long
Private m_lngUnicastPacketsSent As Long
Private m_lngNonunicastPacketsSent As Long
Private m_lngDiscardedOutgoingPackets As Long
Private m_lngOutgoingErrors As Long
Private m_lngOutputQueueLength As Long
Private m_lngInterfaceDescription As String

Public Property Get InterfaceDescription() As String
Attribute InterfaceDescription.VB_Description = "Contains a description of the interface."
    InterfaceDescription = m_lngInterfaceDescription
End Property

Public Property Let InterfaceDescription(strNewValue As String)
    m_lngInterfaceDescription = strNewValue
End Property

Public Property Get OutputQueueLength() As Long
Attribute OutputQueueLength.VB_Description = "Specifies the output queue length."
    OutputQueueLength = m_lngOutputQueueLength
End Property

Public Property Let OutputQueueLength(lngNewValue As Long)
    m_lngOutputQueueLength = lngNewValue
End Property

Public Property Get OutgoingErrors() As Long
Attribute OutgoingErrors.VB_Description = "Specifies the number of outgoing packets that were discarded because of errors."
    OutgoingErrors = m_lngOutgoingErrors
End Property

Public Property Let OutgoingErrors(lngNewValue As Long)
    m_lngOutgoingErrors = lngNewValue
End Property

Public Property Get DiscardedOutgoingPackets() As Long
Attribute DiscardedOutgoingPackets.VB_Description = "Specifies the number of outgoing packets that were discarded even though they did not have errors."
    DiscardedOutgoingPackets = m_lngDiscardedOutgoingPackets
End Property

Public Property Let DiscardedOutgoingPackets(lngNewValue As Long)
    m_lngDiscardedOutgoingPackets = lngNewValue
End Property

Public Property Get NonunicastPacketsSent() As Long
Attribute NonunicastPacketsSent.VB_Description = "Specifies the number of non-unicast packets sent through this interface. Broadcast and multicast packets are included."
    NonunicastPacketsSent = m_lngNonunicastPacketsSent
End Property

Public Property Let NonunicastPacketsSent(lngNewValue As Long)
    m_lngNonunicastPacketsSent = lngNewValue
End Property

Public Property Get UnicastPacketsSent() As Long
Attribute UnicastPacketsSent.VB_Description = "Specifies the number of unicast packets sent through this interface."
    UnicastPacketsSent = m_lngUnicastPacketsSent
End Property

Public Property Let UnicastPacketsSent(lngNewValue As Long)
    m_lngUnicastPacketsSent = lngNewValue
End Property

Public Property Get OctetsSent() As Long
Attribute OctetsSent.VB_Description = "Specifies the number of octets of data sent through this interface."
    OctetsSent = m_lngOctetsSent
End Property

Public Property Let OctetsSent(lngNewValue As Long)
    m_lngOctetsSent = lngNewValue
End Property

Public Property Get UnknownProtocolPackets() As Long
Attribute UnknownProtocolPackets.VB_Description = "Specifies the number of incoming packets that were discarded because the protocol was unknown."
    UnknownProtocolPackets = m_lngUnknownProtocolPackets
End Property

Public Property Let UnknownProtocolPackets(lngNewValue As Long)
    m_lngUnknownProtocolPackets = lngNewValue
End Property

Public Property Get IncomingErrors() As Long
Attribute IncomingErrors.VB_Description = "Specifies the number of incoming packets that were discarded because of errors."
    IncomingErrors = m_lngIncomingErrors
End Property

Public Property Let IncomingErrors(lngNewValue As Long)
    m_lngIncomingErrors = lngNewValue
End Property

Public Property Get DiscardedIncomingPackets() As Long
Attribute DiscardedIncomingPackets.VB_Description = "Specifies the number of incoming packets that were discarded even though they did not have errors."
    DiscardedIncomingPackets = m_lngDiscardedIncomingPackets
End Property

Public Property Let DiscardedIncomingPackets(lngNewValue As Long)
    m_lngDiscardedIncomingPackets = lngNewValue
End Property

Public Property Get NonunicastPacketsReceived() As Long
Attribute NonunicastPacketsReceived.VB_Description = "Specifies the number of non-unicast packets received through this interface. Broadcast and multicast packets are included."
    NonunicastPacketsReceived = m_lngNonunicastPacketsReceived
End Property

Public Property Let NonunicastPacketsReceived(lngNewValue As Long)
    m_lngNonunicastPacketsReceived = lngNewValue
End Property

Public Property Get MaximumTransmissionUnit() As Long
Attribute MaximumTransmissionUnit.VB_Description = "Specifies the Maximum Transmission Unit (MTU)."
    MaximumTransmissionUnit = m_lngMaximumTransmissionUnit
End Property

Public Property Let MaximumTransmissionUnit(lngNewValue As Long)
    m_lngMaximumTransmissionUnit = lngNewValue
End Property

Public Property Get UnicastPacketsReceived() As Long
Attribute UnicastPacketsReceived.VB_Description = "Specifies the number of unicast packets received through this interface."
    UnicastPacketsReceived = m_lngUnicastPacketsReceived
End Property

Public Property Let UnicastPacketsReceived(lngNewValue As Long)
    m_lngUnicastPacketsReceived = lngNewValue
End Property

Public Property Get OctetsReceived() As Long
Attribute OctetsReceived.VB_Description = "Specifies the number of octets of data received through this interface."
    OctetsReceived = m_lngOctetsReceived
End Property

Public Property Let OctetsReceived(lngNewValue As Long)
    m_lngOctetsReceived = lngNewValue
End Property

Public Property Get LastChange() As Date
Attribute LastChange.VB_Description = "Specifies the last time the operational status changed."
    LastChange = m_datLastChange
End Property

Public Property Let LastChange(datNewValue As Date)
    m_datLastChange = datNewValue
End Property

Public Property Get OperationalStatus() As OperationalStates
Attribute OperationalStatus.VB_Description = "Specifies the operational status of the interface."
    OperationalStatus = m_OperationalStatus
End Property

Public Property Let OperationalStatus(NewValue As OperationalStates)
    m_OperationalStatus = NewValue
End Property

Public Property Get AdminStatus() As AdminStatuses
Attribute AdminStatus.VB_Description = "Specifies the interface is administratively enabled or disabled. "
    AdminStatus = m_AdminStatus
End Property

Public Property Let AdminStatus(NewValue As AdminStatuses)
    m_AdminStatus = NewValue
End Property

Public Property Get AdapterAddress() As String
Attribute AdapterAddress.VB_Description = "Specifies the physical address of the adapter for this interface."
    AdapterAddress = m_lngAdapterAddress
End Property

Public Property Let AdapterAddress(strNewValue As String)
    m_lngAdapterAddress = strNewValue
End Property

Public Property Get Speed() As Long
Attribute Speed.VB_Description = "Specifies the speed of the interface in bits per second."
    Speed = m_lngSpeed
End Property

Public Property Let Speed(lngNewValue As Long)
    m_lngSpeed = lngNewValue
End Property

Public Property Get InterfaceType() As InterfaceTypes
Attribute InterfaceType.VB_Description = "Specifies the type of interface."
    InterfaceType = m_InterfaceType
End Property

Public Property Let InterfaceType(NewValue As InterfaceTypes)
    m_InterfaceType = NewValue
End Property

Public Property Get InterfaceIndex() As Long
Attribute InterfaceIndex.VB_Description = "Specifies the index that identifies the interface."
    InterfaceIndex = m_lngInterfaceIndex
End Property

Public Property Let InterfaceIndex(lngNewValue As Long)
    m_lngInterfaceIndex = lngNewValue
End Property

Public Property Get InterfaceName() As String
Attribute InterfaceName.VB_Description = "String that contains the name of the interface. "
    InterfaceName = m_strInterfaceName
End Property

Public Property Let InterfaceName(strNewValue As String)
    m_strInterfaceName = strNewValue
End Property



