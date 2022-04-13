Public Class MSG_GetAttributeListConnectedClass1

    Public CIP_DataRequest As New CIP_ConnectedDataRequestClass1

    Public TagName As String

    'Message Request

    Dim Service1 As Byte = CIP_Services.SC_GET_ATT_LIST

    ' Generate PATH Information

    Public RequestPathSize As Byte = 3 '2 Words
    Public ClassSegment As Byte = &H20
    Public ClassValue As Byte = &H6B
    Public InstanceSegment As Int16 = &H25
    Public InstanceValue As Int16 = &H2BE9

    'Command Specific Data

    Public AttributeCount As Int16
    Public AttributeCountR As Int16

    Public Sub New(SessionHandle As Int32, ConnectionIdentifier As Int32, SequenceCount As Int16)

        CIP_DataRequest.SendUnitData.SessionHandle = SessionHandle
        CIP_DataRequest.SendUnitData.ConnectionIdentifier = ConnectionIdentifier
        CIP_DataRequest.SendUnitData.SequenceCount = SequenceCount

    End Sub

    Public Function Write(ClassSegment As Byte, ClassValue As Byte, InstanceSegment As Int16, Instancevalue As Int16, AttributeList() As Int16) As Byte()

        Dim Mem1 As New System.IO.MemoryStream
        Dim Wrt1 As New System.IO.BinaryWriter(Mem1)

        If InstanceSegment = &H24 Then
            RequestPathSize = 2
        ElseIf InstanceSegment = &H25 Then
            RequestPathSize = 3
        Else
            Return Nothing
        End If

        Wrt1.Write(Service1)
        Wrt1.Write(RequestPathSize)

        Wrt1.Write(ClassSegment)
        Wrt1.Write(ClassValue)

        If RequestPathSize = 2 Then

            Wrt1.Write(CByte(InstanceSegment))
            Wrt1.Write(CByte(Instancevalue))

        ElseIf RequestPathSize = 3 Then

            Wrt1.Write(InstanceSegment)
            Wrt1.Write(Instancevalue)

        End If

        AttributeCount = AttributeList.Length

        Wrt1.Write(AttributeCount)

        For n As Integer = 0 To AttributeCount - 1
            Wrt1.Write(AttributeList(n))
        Next n

        Wrt1.Flush()

        Return CIP_DataRequest.Write(Mem1.ToArray)


    End Function

    Public Function Read(Buffer() As Byte) As Byte()

        Dim Data As Byte() = CIP_DataRequest.Read(Buffer)

        If Data Is Nothing Then Return Nothing

        Dim Mem1 As New System.IO.MemoryStream(Data)
        Dim Rdr1 As New System.IO.BinaryReader(Mem1)

        AttributeCountR = Rdr1.ReadInt16()

        Dim RetLen As Integer = Mem1.Length - Mem1.Position
        Dim RetBuf(RetLen - 1) As Byte
        Array.Copy(Mem1.ToArray, Mem1.Position, RetBuf, 0, RetLen)

        Return RetBuf

    End Function



End Class
