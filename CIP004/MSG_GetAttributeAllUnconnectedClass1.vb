Public Class MSG_GetAttributeAllUnconnectedClass1

    Public CIP_DataRequest As New CIP_UnconnectedDataRequestClass1()

    Public TagName As String

    'Message Request

    Dim Service1 As Byte = CIP_Services.SC_GET_ATT_ALL

    ' Generate PATH Information

    Public RequestPathSize As Byte = 3 '2 Words
    Public ClassSegment As Byte = &H20
    Public ClassValue As Byte = &H6B
    Public InstanceSegment As Int16 = &H25
    Public InstanceValue As Int16 = &H2BE9


    Public Sub New(SessionHandle As Int32, Slot As Byte)

        CIP_DataRequest.SendRRData.SessionHandle = SessionHandle
        CIP_DataRequest.Address = Slot

    End Sub

    Public Function Write(ClassSegment As Byte, ClassValue As Byte, InstanceSegment As Int16, Instancevalue As Int16) As Byte()

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
            Wrt1.Write(CByte(InstanceValue))

        ElseIf RequestPathSize = 3 Then

            Wrt1.Write(InstanceSegment)
            Wrt1.Write(InstanceValue)

        End If

        Wrt1.Flush()

        Return CIP_DataRequest.Write(Mem1.ToArray)


    End Function

    Public Function Read(Buffer() As Byte) As Byte()

      Dim Data As Byte() = CIP_DataRequest.Read(Buffer)

        If Data Is Nothing Then Return Nothing

        Return Data

    End Function



End Class
