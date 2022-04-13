
Public Class CIP_ConnectedDataRequestClass1

    Public SendUnitData As New ENIP_SendUnitDataClass1()

    Public RequestPathSize As Byte = 3 '2 Words
    Public ClassSegment As Byte = &H20
    Public ClassValue As Byte = &HB2
    Public InstanceSegment As Int16 = &H25
    Public InstanceValue As Int16 = &H21

    'Reponse
    Public Service1R As Byte
    Public Pad1R As Byte
    Public StatusR As Byte
    Public AdditionalStatusSizeR As Byte
    Public AdditionalStatusR As Int16

    Public Function Write(MessageRequest As Byte()) As Byte()

        SendUnitData.InterfaceHandle = 0 'CIP
        Return (SendUnitData.Write(MessageRequest))

    End Function

    Public Function Read(Buffer() As Byte) As Byte()

        Dim Buffer1 As Byte() = SendUnitData.Read(Buffer)
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

        If StatusR <> 0 Then Return Nothing

        Dim RetLen As Integer = Mem1.Length - Mem1.Position
        Dim RetBuf(RetLen - 1) As Byte
        Array.Copy(Mem1.ToArray, Mem1.Position, RetBuf, 0, RetLen)

        Return RetBuf

    End Function

End Class
