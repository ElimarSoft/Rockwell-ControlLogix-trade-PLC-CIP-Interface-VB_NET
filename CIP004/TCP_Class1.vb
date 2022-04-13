Public Class TCP_Class1

    Public TCP1 As New System.Net.Sockets.TcpClient()
    Private St1 As System.IO.Stream
    Public RecBuf(4095) As Byte
    Private RecData As Byte()

    Private SessionHandle As Int32
    Private ConnectionIdentifier As Int32
    Private SequenceCount As Int16

    Public Slot As Byte = 0
    Private LogSt1 As System.IO.FileStream


    Public Sub New()

        Try

            TCP1.Connect(IP_ADD, 44818)
            St1 = TCP1.GetStream

        Catch e As Exception

            MsgBox(e.Message)
            Application.Exit()
            End

        End Try
    End Sub

    Public Sub Connect()
        RegSession()

    End Sub

    Public Sub RegSession()

        Dim RegSession As New ENIP_RegisterSessionClass1
        SendRequest(RegSession.Write)
        SessionHandle = RegSession.Read(ReadAnswer)

    End Sub


    Public Sub ForwardOpen()

        Dim ForwardOpen As New CIP_ForwardOpenClass1(SessionHandle, Slot)
        SendRequest(ForwardOpen.Write)
        ConnectionIdentifier = ForwardOpen.Read(ReadAnswer())

    End Sub

    Public Sub ForwardClose()

        Dim ForwardClose As New CIP_ForwardCloseClass1(SessionHandle, Slot)
        SendRequest(ForwardClose.Write)
        ForwardClose.Read(ReadAnswer())

    End Sub

    Public Sub Send()

        Dim DataValW1(1023) As Object

        LogSt1 = New System.IO.FileStream(FilePath + "Out.Dat", IO.FileMode.Create, IO.FileAccess.Write)

        DataValW1(0) = CInt(546)

        For n As Integer = 0 To 127
            DataValW1(n) = CInt(8900) + n * 160
        Next n

        ForwardOpen()

        Dim Data As Byte() = GetAttributeList(&H20, &HAC, &H24, &H1, {1, 2, 3, 4, 10}, True)

        DisplayData(Data)


        'GetAttributeAll(&H20, &HAC, &H24, &H1, True)

        ReadTagMult("TAG4", DataTypes.DINT, 1, True)


        'GetAttributeAll(&H20, &H6C, &H25, &H8B, False)
        'GetAttributeAll(&H20, &H6B, &H25, &H2BE9, False)

        'GetAttributeList(&H20, &H6B, &H25, &H2BE9, {2, 3, 4, 5, 6}, True)
        'GetAttributeList(&H20, &H6B, &H25, &H2BE9, {1, 2, 3, 4, 5, 6}, False)


        'GetAttributeList(&H20, &H6B, &H25, &H2BE9, {2}, True)
        'GetAttributeList(&H20, &H6B, &H25, &H2BE9, {3}, True)
        'GetAttributeList(&H20, &H6B, &H25, &H2BE9, {4}, True)
        'GetAttributeList(&H20, &H6B, &H25, &H2BE9, {5}, True)
        'GetAttributeList(&H20, &H6B, &H25, &H2BE9, {6}, True)
        'GetAttributeList(&H20, &H6B, &H25, &H2BE9, {7}, True)
        'GetAttributeList(&H20, &H6B, &H25, &H2BE9, {8}, True)
        'GetAttributeList(&H20, &H6B, &H25, &H2BE9, {9}, True)
        'GetAttributeList(&H20, &H6B, &H25, &H2BE9, {10}, True)
        'GetAttributeList(&H20, &H6B, &H25, &H2BE9, {11}, True)
        'GetAttributeList(&H20, &H6B, &H25, &H2BE9, {12}, True)

        'GetAttributeList(&H20, &H6B, &H25, &H2BE9, {13}, True)
        'GetAttributeList(&H20, &H6B, &H25, &H2BE9, {14}, True)
        'GetAttributeList(&H20, &H6B, &H25, &H2BE9, {15}, True)
        'GetAttributeList(&H20, &H6B, &H25, &H2BE9, {16}, True)
        'GetAttributeList(&H20, &H6B, &H25, &H2BE9, {17}, True)
        'GetAttributeList(&H20, &H6B, &H25, &H2BE9, {18}, True)
        'GetAttributeList(&H20, &H6B, &H25, &H2BE9, {19}, True)
        'GetAttributeList(&H20, &H6B, &H25, &H2BE9, {20}, True)
        'GetAttributeList(&H20, &H6B, &H25, &H2BE9, {21}, True)
        'GetAttributeList(&H20, &H6B, &H25, &H2BE9, {22}, True)
        'GetAttributeList(&H20, &H6B, &H25, &H2BE9, {23}, True)
        'GetAttributeList(&H20, &H6B, &H25, &H2BE9, {24}, True)

        'GetAttributeList(&H20, &H6C, &H25, &H8B, {1, 2, 3, 5, 6, 7, 8, 9, 10, 11, 12}, False)

        'Dim DataValR1 As Object = ReadTag("TAG4", DataTypes.INT, 1, True)
        'Dim DataValR2 As Object = ReadTag("ARRAY1", DataTypes.DINT, 16, True)

        'ReadTag("TAG4", DataTypes.DINT, 1, True)

        'WriteTag("ARRAY1", DataTypes.DINT, 16, DataValW1, False)
        'WriteTag("ARRAY2", DataTypes.DINT, 16, DataValW1, True)

        'Dim Data As New CIP_MultipleServicePacketClass1(SessionHandle, ConnectionIdentifier, SequenceCount, Slot)
        'SendRequest(Data.Write)
        'SequenceCount += 1
        'Data.Read(ReadAnswer)


        'GetTagNames(0, {4, 2, 7, 8, 1})
        'GetTagNames(&H2583, {4, 2, 7, 8, 1})
        'GetTagNames(&H3D81, {4, 2, 7, 8, 1})
        'GetTagNames(&H5D14, {4, 2, 7, 8, 1})
        'GetTagNames(&H7738, {4, 2, 7, 8, 1})
        'GetTagNames(&H9146, {4, 2, 7, 8, 1})
        'GetTagNames(&HB7D3, {4, 2, 7, 8, 1})
        'GetTagNames(&HD672, {4, 2, 7, 8, 1})
        'GetTagNames(&HF43B, {4, 2, 7, 8, 1})

        ForwardClose()

        LogSt1.Close()


        Dim Stop1 As Byte = 0
    End Sub
    Public Sub Close()

        ForwardClose()

    End Sub

    Public Function GetTagNames(InstanceValue As UInt16, AttributeList As Int16()) As Byte()

        Dim Data As New MSG_GetTagNamesUnconnectedClass1(SessionHandle, Slot)
        SendRequest(Data.Write(InstanceValue, AttributeList))
        Dim Result As Byte() = Data.Read(ReadAnswer())
        LogSt1.Write(Result, 0, Result.Length)
        Return Result

    End Function




    Public Function GetAttributeAll(ClassSegment As Byte, ClassValue As Byte, InstanceSegment As Int16, InstanceValue As Int16, Connected As Boolean)

        If Connected Then

            Dim Data As New MSG_GetAttributeAllConnectedClass1(SessionHandle, ConnectionIdentifier, SequenceCount)
            SendRequest(Data.Write(ClassSegment, ClassValue, InstanceSegment, InstanceValue))
            SequenceCount += 1
            Return Data.Read(ReadAnswer())

        Else

            Dim Data As New MSG_GetAttributeAllUnconnectedClass1(SessionHandle, Slot)
            SendRequest(Data.Write(ClassSegment, ClassValue, InstanceSegment, InstanceValue))
            Return Data.Read(ReadAnswer())

        End If

    End Function


    Public Function GetAttributeList(ClassSegment As Byte, ClassValue As Byte, InstanceSegment As Int16, InstanceValue As Int16, AttributeList() As Int16, Connected As Boolean) As Byte()


        If Connected Then

            Dim Data As New MSG_GetAttributeListConnectedClass1(SessionHandle, ConnectionIdentifier, SequenceCount)
            SendRequest(Data.Write(ClassSegment, ClassValue, InstanceSegment, InstanceValue, AttributeList))
            SequenceCount += 1
            Return Data.Read(ReadAnswer())

        Else
            Dim Data As New MSG_GetAttributeListUnconnectedClass1(SessionHandle, Slot)
            SendRequest(Data.Write(ClassSegment, ClassValue, InstanceSegment, InstanceValue, AttributeList))
            Return Data.Read(ReadAnswer())
        End If

    End Function

    Public Function WriteTag(Tag As String, Datatype As DataTypes, ReqElemCount1 As Int16, Dataval As Object, Connected As Boolean) As Boolean


        If Connected Then

            Dim Data As New RA_ConnectedWriteTagClass1(Tag, Datatype, ReqElemCount1, Dataval, SessionHandle, ConnectionIdentifier, SequenceCount)
            SendRequest(Data.Write)
            SequenceCount += 1
            Return Data.Read(ReadAnswer())

        Else

            Dim Data As New RA_UnconnectedWriteTagClass1(Tag, Datatype, ReqElemCount1, Dataval, SessionHandle, Slot)
            SendRequest(Data.Write)
            Return Data.Read(ReadAnswer())

        End If


    End Function

    Public Function ReadTagMult(Tag As String, Datatype As DataTypes, ReqElemCount1 As Int16, Connected As Boolean) As Object()

        If Connected Then

            Dim Data As New RA_ConnectedReadTagMultClass1(Tag, ReqElemCount1, SessionHandle, ConnectionIdentifier, SequenceCount)
            SendRequest(Data.Write)
            SequenceCount += 1
            Return Data.Read(ReadAnswer())

        Else

            Return Nothing

        End If


    End Function
    Public Function ReadTag(Tag As String, Datatype As DataTypes, ReqElemCount1 As Int16, Connected As Boolean) As Object()

        If Connected Then

            Dim Data As New RA_ConnectedReadTagMultClass1(Tag, ReqElemCount1, SessionHandle, ConnectionIdentifier, SequenceCount)
            SendRequest(Data.Write)
            SequenceCount += 1
            Return Data.Read(ReadAnswer())

        Else

            Return Nothing

        End If

    End Function
    Private Function ReadAnswer() As Byte()

        Dim Timeout As Integer = 0
        Dim ByteCount As Integer = 0

        Do
            System.Threading.Thread.Sleep(10)

            ByteCount = St1.Read(RecBuf, 0, RecBuf.Length)
            If ByteCount > 0 Then Exit Do

            Timeout += 1
            If Timeout > 10 Then Exit Do

        Loop

        Dim RecData(ByteCount - 1) As Byte

        Array.Copy(RecBuf, RecData, ByteCount)

        Return RecData

    End Function


    Public Sub SendRequest(Data As Byte())

        St1.Write(Data, 0, Data.Length)

    End Sub

    Public Sub DisplayData(Data As Byte())

        Dim Mem1 As New System.IO.MemoryStream(Data)
        Dim Rdr1 As New System.IO.BinaryReader(Mem1)
        Dim Str1 As String = vbNullString

        For n As Integer = 0 To (Data.Length / 2) - 1
            Dim Val1 As Int16 = Rdr1.ReadInt16()
            Str1 += Right("0000" + Hex(Val1), 4) + " | "
        Next n

        Form1.ListBox1.Items.Add(Str1)


    End Sub


End Class
