Public Class RA_UnconnectedWriteTagClass1

    Public CIP_DataRequest As New CIP_UnconnectedDataRequestClass1()

    Public TagName As String

    'Message Request

    Dim Service1 As Byte = RA_Services.CIP_WriteData
    Dim RequestPathSize1 As Byte

    ' Generate TAG Information

    Const TagPrefix As Byte = &H91
    Private TagSize As Byte
    Private TagSizePadded As Byte

    Private DataType As Int16
    Private DataCount As Int16
    Private DataVal() As Object

    Public Sub New(TagName As String, DataType As Int16, DataCount As Int16, DataVal() As Object, SessionHandle As Int32, Slot As Byte)

        Me.TagName = TagName

        CIP_DataRequest.SendRRData.SessionHandle = SessionHandle
        CIP_DataRequest.Address = Slot

        Me.DataType = DataType
        Me.DataCount = DataCount
        Me.DataVal = DataVal

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


        Select Case DataType

            Case DataTypes.DINT

                Wrt1.Write(DataType)
                Wrt1.Write(DataCount)

                For n = 0 To DataCount - 1
                    Wrt1.Write(CInt(DataVal(n)))
                Next

            Case DataTypes.INT

                Wrt1.Write(DataType)
                Wrt1.Write(DataCount)

                For n = 0 To DataCount - 1
                    Wrt1.Write(CShort(DataVal(n)))
                Next

            Case DataTypes.SINT

                Wrt1.Write(DataType)
                Wrt1.Write(DataCount)

                For n = 0 To DataCount - 1
                    Wrt1.Write(CByte(DataVal(n)))
                Next

            Case Else

        End Select

        Wrt1.Flush()

        Return CIP_DataRequest.Write(Mem1.ToArray)


    End Function

    Public Function Read(Buffer() As Byte) As Boolean

        Dim Data As Byte() = CIP_DataRequest.Read(Buffer)
        Return (CIP_DataRequest.StatusR <> 0) Or (CIP_DataRequest.SendRRData.StatusR <> 0)

    End Function



End Class
