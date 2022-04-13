Public Class CIP_ForwardOpenClass1

    Public SendRRData As New ENIP_SendRRDataClass1()

    Public Service As Byte = CIP_Services.SC_FWD_OPEN
    Public RequestPathSize As Byte = 2 '2 Words
    Public ClassSegment As Byte = &H20
    Public ClassValue As Byte = 6
    Public InstanceSegment As Byte = &H24
    Public InstanceValue As Byte = 1

    'Command Specific Data

    Public PriorityTimeTick As Byte = 7
    Public TimeOutTicks As Byte = 233
    Public O_T_ConnID As Int32 = &H2
    Public T_O_ConnID As Int32 = &H1
    Public ConnectionSerialNumber As Int16 = 2 '2
    Public VendorID As Int16 = &H1
    Public OriginatorSerialNumber As Int32 = &H27B01B0
    Public ConnectionTimeoutMultiplier As Byte = 2
    Public ReservedData1 As Byte = 0
    Public ReservedData2 As Byte = 0
    Public ReservedData3 As Byte = 0
    Public O_T_RPI As Int32 = &H1E8480
    Public O_T_NetworkConnectionParam As Int16 = &H43F4
    Public T_O_RPI As Int32 = &H1E8480
    Public T_O_NetworkConnectionParam As Int16 = &H43F4
    Public TransportTypeTrigger As Byte = &HA3

    'Connection Path
    Public ConnectionPathSize As Byte = 3 'Words 4 If ConnectionPointSegment1
    Public Port As Byte = 1
    Public Address As Byte

    Public ClassSegment1 As Byte = &H20
    Public ClassValue1 As Byte = 2 'Message Router
    Public InstanceSegment1 As Byte = &H24
    Public InstanceValue1 As Byte = 1
    Public ConnectionPointSegment1 As Byte = &H2C
    Public ConnectionPoint1 As Byte = 1

    'Reponse

    Public Service1R As Byte = 1 'GetAttributeAll
    Public Pad1R As Byte
    Public StatusR As Byte
    Public AdditionalStatusSizeR As Byte
    Public AdditionalStatusR As Int16

    Public O_T_ConnID_R As Int32
    Public T_O_ConnID_R As Int32
    Public ConnectionSerialNumberR As Int16
    Public VendorID_R As Int16
    Public OriginatorSerialNumberR As Int32

    Public O_T_API As Int32
    Public T_O_API As Int32

    Public ApplicationReplaySizeR As Byte
    Public ReserveR As Byte
    Public Sub New(SessionHandle As Int32, Address As Byte)

        SendRRData.SessionHandle = SessionHandle
        Me.Address = Address

    End Sub

    Public Function Write() As Byte()

        Dim Mem1 As New System.IO.MemoryStream
        Dim Wrt1 As New System.IO.BinaryWriter(Mem1)


        Wrt1.Write(Service)
        Wrt1.Write(RequestPathSize)
        Wrt1.Write(ClassSegment)
        Wrt1.Write(ClassValue)
        Wrt1.Write(InstanceSegment)
        Wrt1.Write(InstanceValue)
        Wrt1.Write(PriorityTimeTick)
        Wrt1.Write(TimeOutTicks)

        Wrt1.Write(O_T_ConnID)
        Wrt1.Write(T_O_ConnID)
        Wrt1.Write(ConnectionSerialNumber)
        Wrt1.Write(VendorID)
        Wrt1.Write(OriginatorSerialNumber)
        Wrt1.Write(ConnectionTimeoutMultiplier)
        Wrt1.Write(ReservedData1)
        Wrt1.Write(ReservedData2)
        Wrt1.Write(ReservedData3)
        Wrt1.Write(O_T_RPI)
        Wrt1.Write(O_T_NetworkConnectionParam)
        Wrt1.Write(T_O_RPI)
        Wrt1.Write(T_O_NetworkConnectionParam)
        Wrt1.Write(TransportTypeTrigger)

        Wrt1.Write(ConnectionPathSize)
        Wrt1.Write(Port)
        Wrt1.Write(Address)

        Wrt1.Write(ClassSegment1)
        Wrt1.Write(ClassValue1)
        Wrt1.Write(InstanceSegment1)
        Wrt1.Write(InstanceValue1)

        Wrt1.Flush()

        Return SendRRData.Write(Mem1.ToArray)

    End Function

    Public Function Read(Buffer() As Byte) As Int32

        Dim Data As Byte() = SendRRData.Read(Buffer)
        If IsNothing(Data) Then Return Nothing
        'If Buffer1.Length <> 30 Then Return Nothing

        Dim Mem1 As New System.IO.MemoryStream(Data)
        Dim Rdr1 As New System.IO.BinaryReader(Mem1)

        Service1R = Rdr1.ReadByte()
        Pad1R = Rdr1.ReadByte()

        StatusR = Rdr1.ReadByte()
        AdditionalStatusSizeR = Rdr1.ReadByte()

        If AdditionalStatusSizeR = 1 Then
            AdditionalStatusR = Rdr1.ReadInt16
        End If

        If StatusR <> 0 Then Return Nothing

        O_T_ConnID_R = Rdr1.ReadInt32
        T_O_ConnID_R = Rdr1.ReadInt32
        ConnectionSerialNumberR = Rdr1.ReadInt16
        VendorID_R = Rdr1.ReadInt16
        OriginatorSerialNumberR = Rdr1.ReadInt32

        O_T_API = Rdr1.ReadInt32
        T_O_API = Rdr1.ReadInt32

        ApplicationReplaySizeR = Rdr1.ReadByte
        ReserveR = Rdr1.ReadByte

        Return O_T_ConnID_R

    End Function

End Class
