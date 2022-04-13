
Public Class CIP_MultipleServicePacketClass1

    Public SendUnitData As New ENIP_SendUnitDataClass1()

    Public Service As Byte = CIP_Services.SC_MULT_SERV_PACK

    Public RequestPathSize As Byte = 2 '2 Words
    Public ClassSegment As Byte = &H20
    Public ClassValue As Byte = &H2
    Public InstanceSegment As Byte = &H24
    Public InstanceValue As Byte = &H1

    'Route Path 

    'Public RoutePathSize As Byte = 1 'Words
    'Public Reserve As Byte = 0 'Pad for Even

    'Public Port As Byte = 1
    'Public Address As Byte

    'Reponse
    Public Service1R As Byte = CIP_Services.SC_MULT_SERV_PACK
    Public Pad1R As Byte
    Public StatusR As Byte
    Public AdditionalStatusSizeR As Byte
    Public AdditionalStatusR As Int16
    Public NumberOfServices As Int16 = 4
    Public Address As Byte

    Public ServicePacket(NumberOfServices - 1) As ServicePacketClass1

    Public Sub New(SessionHandle As Int32, ConnectionIdentifier As Int32, SequenceCount As Int16, Slot As Byte)

        SendUnitData.SessionHandle = SessionHandle
        SendUnitData.ConnectionIdentifier = ConnectionIdentifier
        SendUnitData.SequenceCount = SequenceCount
        Me.Address = Slot

        For n As Integer = 0 To NumberOfServices - 1
            ServicePacket(n) = New ServicePacketClass1
        Next

        ServicePacket(0).SndData = GetAttributeList()
        ServicePacket(1).SndData = ReadTag("INTARRAY1", 16)
        ServicePacket(2).SndData = ReadTag("TAG5", 1)
        ServicePacket(3).SndData = ReadTag("TAG6", 1)

        ServicePacket(0).SndOffset = 2 + (2 * NumberOfServices)

        For n As Integer = 1 To NumberOfServices - 1
            ServicePacket(n).SndOffset = ServicePacket(n - 1).SndOffset + ServicePacket(n - 1).SndData.Length
        Next n

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

        Wrt1.Write(NumberOfServices)

        For n As Integer = 0 To NumberOfServices - 1
            Wrt1.Write(ServicePacket(n).SndOffset)
        Next

        For n As Integer = 0 To NumberOfServices - 1
            Wrt1.Write(ServicePacket(n).SndData)
        Next n

        Wrt1.Flush()

        SendUnitData.InterfaceHandle = 0 'CIP
        Return (SendUnitData.Write(Mem1.ToArray))

    End Function

    Public Function Read(Buffer() As Byte) As Boolean

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

        If StatusR <> 0 Then Return False

        Dim NumberOfReplies As Int16 = Rdr1.ReadInt16()

        For n As Integer = 0 To NumberOfReplies - 1
            ServicePacket(n).RecOffset = Rdr1.ReadInt16()
        Next

        Dim OffsetPos As Integer = Mem1.Position - ServicePacket(0).RecOffset


        For n As Integer = 0 To NumberOfReplies - 2
            ServicePacket(n).RecLength = ServicePacket(n + 1).RecOffset - ServicePacket(n).RecOffset
        Next
        ServicePacket(NumberOfReplies - 1).RecLength = Buffer1.Length - OffsetPos - ServicePacket(NumberOfReplies - 1).RecOffset

        For n As Integer = 0 To NumberOfReplies - 1
            ServicePacket(n).RecData = Array.CreateInstance(GetType(Byte), ServicePacket(n).RecLength)
            Array.Copy(Buffer1, OffsetPos + ServicePacket(n).RecOffset, ServicePacket(n).RecData, 0, ServicePacket(n).RecLength)
        Next

        Return True

    End Function

    Public Function ReadTag(TagName As String, RequestDataCount1 As Int16) As Byte()

        Dim Service1 As Byte = CIP_Services.SC_UNCONECTED_SEND
        Dim RequestPathSize1 As Byte

        Const TagPrefix As Byte = &H91
        Dim TagSize As Byte
        Dim TagSizePadded As Byte
        Dim Pad1 As Int32

        Dim Mem1 As New System.IO.MemoryStream
        Dim Wrt1 As New System.IO.BinaryWriter(Mem1)

        TagSize = TagName.Length
        TagSizePadded = TagName.Length

        If (TagSize Mod 2) > 0 Then
            TagName += vbNullChar
            TagSizePadded += 1
        End If

        Dim RequestPath1 As Char() = TagName.ToCharArray
        RequestPathSize1 = (TagSizePadded + 2) / 2

        Wrt1.Write(Service1)
        Wrt1.Write(RequestPathSize1)
        Wrt1.Write(TagPrefix)
        Wrt1.Write(TagSize)
        Wrt1.Write(RequestPath1)
        Wrt1.Write(RequestDataCount1)

        Wrt1.Write(Pad1)

        Wrt1.Flush()

        Return Mem1.ToArray


    End Function

    Public Function GetAttributeList()

        Dim Service As Byte = CIP_Services.SC_GET_ATT_LIST

        Dim RequestPathSize As Byte = 2 '2 Words
        Dim ClassSegment As Byte = &H20
        Dim ClassValue As Byte = &HAC '172
        Dim InstanceSegment As Byte = &H24
        Dim InstanceValue As Byte = 1

        Dim AttributeCount As Int16 = 2
        Dim AttributeList() As Int16 = {1, 3}

        Dim Mem1 As New System.IO.MemoryStream
        Dim Wrt1 As New System.IO.BinaryWriter(Mem1)

        Wrt1.Write(Service)
        Wrt1.Write(RequestPathSize)
        Wrt1.Write(ClassSegment)
        Wrt1.Write(ClassValue)
        Wrt1.Write(InstanceSegment)
        Wrt1.Write(InstanceValue)
        Wrt1.Write(AttributeCount)

        For n As Integer = 0 To AttributeCount - 1
            Wrt1.Write(AttributeList(n))
        Next n

        Wrt1.Flush()

        Return Mem1.ToArray

    End Function

End Class
