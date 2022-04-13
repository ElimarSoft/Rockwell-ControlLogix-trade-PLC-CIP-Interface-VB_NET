
Public Class CIP_UnconnectedDataRequestClass1

    Public SendRRData As New ENIP_SendRRDataClass1()

    Public Service As Byte = CIP_Services.SC_UNCONECTED_SEND

    Public RequestPathSize As Byte = 2 '2 Words
    Public ClassSegment As Byte = &H20
    Public ClassValue As Byte = 6
    Public InstanceSegment As Byte = &H24
    Public InstanceValue As Byte = 1

    'Command Specific Data

    Public PriorityTimeTick As Byte = 7
    Public TimeOutTicks As Byte = 233

    'Public ActualTimeOut As Int16 = CDbl(PriorityTimeTick) * CDbl(TimeOutTicks) * 18.2

    'Message Request Comes from the call

    'Route Path 

    Public RoutePathSize As Byte = 1 'Words
    Public Reserve As Byte = 0 'Pad for Even

    Public Port As Byte = 1
    Public Address As Byte

    'Reponse
    Public Service1R As Byte
    Public Pad1R As Byte
    Public StatusR As Byte
    Public AdditionalStatusSizeR As Byte
    Public AdditionalStatusR As Int16

    Public Function Write(MessageRequest As Byte()) As Byte()
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

        Wrt1.Write(CShort(MessageRequest.Length))
        Wrt1.Write(MessageRequest)

        'Route
        Wrt1.Write(RoutePathSize)
        Wrt1.Write(Reserve)
        Wrt1.Write(Port)
        Wrt1.Write(Address)

        Wrt1.Flush()

        SendRRData.InterfaceHandle = 0 'CIP
        Return SendRRData.Write(Mem1.ToArray)

    End Function

    Public Function Read(Buffer() As Byte) As Byte()

        Dim Buffer1 As Byte() = SendRRData.Read(Buffer)
        If IsNothing(Buffer1) Then Return Nothing

        Dim Mem1 As New System.IO.MemoryStream(Buffer1)
        Dim Rdr1 As New System.IO.BinaryReader(Mem1)

        Service1R = Rdr1.ReadByte()
        Pad1R = Rdr1.ReadByte()

        StatusR = Rdr1.ReadByte()
        AdditionalStatusSizeR = Rdr1.ReadByte()

        If AdditionalStatusSizeR = 1 Then
            AdditionalStatusR = Rdr1.ReadInt16
        End If

        If Not ((StatusR = 0) Or (StatusR = 6)) Then Return Nothing ' Status Ok or Partial Transfer

        Dim RetLen As Integer = Mem1.Length - Mem1.Position
        Dim RetBuf(RetLen - 1) As Byte
        Array.Copy(Mem1.ToArray, Mem1.Position, RetBuf, 0, RetLen)

        Return RetBuf

    End Function

End Class
