Public Class CIP_ForwardCloseClass1

    Public SendRRData As New ENIP_SendRRDataClass1()

    Public Service As Byte = CIP_Services.SC_FWD_CLOSE
    Public RequestPathSize As Byte = 2 '2 Words
    Public ClassSegment As Byte = &H20
    Public ClassValue As Byte = 6
    Public InstanceSegment As Byte = &H24
    Public InstanceValue As Byte = 1

    'Command Specific Data

    Public PriorityTimeTick As Byte = 7
    Public TimeOutTicks As Byte = 233
    Public ConnectionSerialNumber As Int16 = 2 '2
    Public VendorID As Int16 = &H1
    Public OriginatorSerialNumber As Int32 = &H27B01B0

    'Connection Path
    Public ConnectionPathSize As Byte = 3 'Words 4 If ConnectionPointSegment1
    Public Reserved1 As Byte
    Public Port As Byte = 1
    Public Address As Byte

    Public ClassSegment1 As Byte = &H20
    Public ClassValue1 As Byte = 2 'Message Router
    Public InstanceSegment1 As Byte = &H24
    Public InstanceValue1 As Byte = 1

    'Reponse

    Public Service1R As Byte = 1 'GetAttributeAll
    Public Pad1R As Byte
    Public StatusR As Byte
    Public AdditionalStatusSizeR As Byte
    Public AdditionalStatusR As Int16

     Public ConnectionSerialNumberR As Int16
    Public VendorID_R As Int16
    Public OriginatorSerialNumberR As Int32

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

        Wrt1.Write(ConnectionSerialNumber)
        Wrt1.Write(VendorID)
        Wrt1.Write(OriginatorSerialNumber)
 
        Wrt1.Write(ConnectionPathSize)
        Wrt1.Write(Reserved1)
        Wrt1.Write(Port)
        Wrt1.Write(Address)

        Wrt1.Write(ClassSegment1)
        Wrt1.Write(ClassValue1)
        Wrt1.Write(InstanceSegment1)
        Wrt1.Write(InstanceValue1)

        Wrt1.Flush()

        Return SendRRData.Write(Mem1.ToArray)

    End Function

    Public Function Read(Buffer() As Byte) As Byte

        Dim Data As Byte() = SendRRData.Read(Buffer)
        If IsNothing(Data) Then Return Nothing
  
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

        ConnectionSerialNumberR = Rdr1.ReadInt16
        VendorID_R = Rdr1.ReadInt16
        OriginatorSerialNumberR = Rdr1.ReadInt32

        ApplicationReplaySizeR = Rdr1.ReadByte
        ReserveR = Rdr1.ReadByte

        Return StatusR

    End Function

End Class
