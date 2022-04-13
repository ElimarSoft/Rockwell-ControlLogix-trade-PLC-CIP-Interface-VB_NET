Public Class RA_ConnectedReadTagMultClass1

    Public CIP_DataRequest As New CIP_ConnectedDataRequestClass1

    Public TagName As String

    'Message Request

    Dim Service1 As Byte = RA_Services.CIP_ReadDataFragmented
    Dim RequestPathSize1 As Byte

    ' Generate TAG Information

    Const TagPrefix As Byte = &H91
    Dim TagSize As Byte
    Dim TagSizePadded As Byte
    Dim RequestDataCount1 As Int16
    Dim Offset As Int16 = 0
    Public Sub New(TagName As String, ReqElemCount1 As Int16, SessionHandle As Int32, ConnectionIdentifier As Int32, SequenceCount As Int16)

        Me.TagName = TagName
        Me.RequestDataCount1 = ReqElemCount1

        CIP_DataRequest.SendUnitData.SessionHandle = SessionHandle
        CIP_DataRequest.SendUnitData.ConnectionIdentifier = ConnectionIdentifier
        CIP_DataRequest.SendUnitData.SequenceCount = SequenceCount

    End Sub

    Public Function Write() As Byte()

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
        Wrt1.Write(Offset)

        Wrt1.Flush()

        Return CIP_DataRequest.Write(Mem1.ToArray)


    End Function

    Public Function Read(Buffer() As Byte) As Object()

        Dim Data As Byte() = CIP_DataRequest.Read(Buffer)

        If (CIP_DataRequest.Service1R And &H7F) <> Service1 Then Return Nothing

        If Data Is Nothing Then Return Nothing

        Dim Mem1 As New System.IO.MemoryStream(Data)
        Dim Rdr1 As New System.IO.BinaryReader(Mem1)

        Dim VarType As Int16 = Rdr1.ReadInt16()
        Dim Values(RequestDataCount1 - 1) As Object

        Select Case VarType

            Case DataTypes.DINT

                If (Data.Length <> ((RequestDataCount1 * 4) + 2)) Then Return Nothing

                For n As Integer = 0 To RequestDataCount1 - 1
                    Values(n) = Rdr1.ReadInt32
                Next n

                Return Values

            Case DataTypes.INT

                If (Data.Length <> ((RequestDataCount1 * 2) + 2)) Then Return Nothing

                For n As Integer = 0 To RequestDataCount1 - 1
                    Values(n) = Rdr1.ReadInt16
                Next n

                Return Values

            Case DataTypes.SINT

                If (Data.Length <> ((RequestDataCount1) + 2)) Then Return Nothing

                For n As Integer = 0 To RequestDataCount1 - 1
                    Values(n) = Rdr1.ReadByte
                Next n

                Return Values

            Case Else
                Return Nothing

        End Select

    End Function



End Class
