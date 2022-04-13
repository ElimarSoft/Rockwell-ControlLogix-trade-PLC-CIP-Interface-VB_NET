Public Class MSG_GetTagNamesUnconnectedClass1

    Public CIP_DataRequest As New CIP_UnconnectedDataRequestClass1()

    Public TagName As String

    'Message Request

    Dim Service1 As Byte = &H55 'Undocummented

    ' Generate PATH Information

    Public RequestPathSize As Byte = 3 '2 Words
    Public ClassSegment As Byte = &H20
    Public ClassValue As Byte = &H6B
    Public InstanceSegment As Int16 = &H25
    Public InstanceValue As Int16

    'Command Specific Data

  
    Public Sub New(SessionHandle As Int32, Slot As Byte)

        CIP_DataRequest.SendRRData.SessionHandle = SessionHandle
        CIP_DataRequest.Address = Slot

    End Sub

    Public Function Write(Instancevalue As UInt16, AttributeList() As Int16) As Byte()

        Dim Mem1 As New System.IO.MemoryStream
        Dim Wrt1 As New System.IO.BinaryWriter(Mem1)

        Wrt1.Write(Service1)
        Wrt1.Write(RequestPathSize)

        Wrt1.Write(ClassSegment)
        Wrt1.Write(ClassValue)

        Wrt1.Write(InstanceSegment)
        Wrt1.Write(Instancevalue)

        For n As Integer = 0 To AttributeList.Length - 1
            Wrt1.Write(AttributeList(n))
        Next n

        Wrt1.Flush()

        Return CIP_DataRequest.Write(Mem1.ToArray)


    End Function

    Public Function Read(Buffer() As Byte) As Byte()

        Dim Data As Byte() = CIP_DataRequest.Read(Buffer)

        If Data Is Nothing Then Return Nothing

        Return Data

    End Function



End Class
