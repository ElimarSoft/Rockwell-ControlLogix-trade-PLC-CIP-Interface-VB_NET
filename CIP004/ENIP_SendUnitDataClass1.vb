Public Class ENIP_SendUnitDataClass1


    Private Command As ENIP_Function_Codes 'Int16
    Private Status As ENIP_Status_Codes 'Int32
    Public Length As Int16
    Public SessionHandle As Int32
    Public SenderContext As Int64 = &H69B5B80043323BDE 'Sequence number question - answer
    Public Slot As Byte = 0
    Public Options As Int32
    Public InterfaceHandle As Int32

    'From here Length is calculated

    'Dim InterfaceHandle As Int32 = &HDF9768 '0 For CIP
    Public TimeOut As Int16 = 32
    Public ItemCount As Int16 = 2

    Public TypeID1 As Int16 = &HA1 'Connected Address Item
    Public Length1 As Int16 = 4
    Public ConnectionIdentifier As Int32

    Public TypeID2 As Int16 = &HB1 'Connected Data Item
    Public Length2 As Int16 = 0
    Public SequenceCount As Int16

    'Received Data

    Private CommandR As ENIP_Function_Codes 'Int16
    Public StatusR As Int32
    Public LengthR As Int16
    Public SessionHandleR As Int32
    Public SenderContextR As Int64
    Public OptionsR As Int32

    Public InterfaceHandleR As Int32

    Public TimeOutR As Int16
    Public ItemCountR As Int16

    Public TypeID1R As Int16
    Public Length1R As Int16
    Public ConnectionIdentifierR As Int32

    Public TypeID2R As Int16
    Public Length2R As Int16
    Public SequenceCountR As Int16


    Public Sub New()

        Command = ENIP_Function_Codes.SEND_UNIT_DATA
        Status = ENIP_Status_Codes.SUCCESS
        SenderContext += 200
        Options = 0

    End Sub

    Public Function Write(SendData As Byte()) As Byte()

        Dim Mem1 As New System.IO.MemoryStream
        Dim Wrt1 As New System.IO.BinaryWriter(Mem1)

        Length2 = SendData.Length + 2
        Length = 16 + Length1 + Length2

        Wrt1.Write(Command)
        Wrt1.Write(Length)
        Wrt1.Write(SessionHandle)
        Wrt1.Write(Status)
        Wrt1.Write(SenderContext)
        Wrt1.Write(Options)

        Wrt1.Write(InterfaceHandle)
        Wrt1.Write(TimeOut)
        Wrt1.Write(ItemCount)

        Wrt1.Write(TypeID1)
        Wrt1.Write(Length1)
        Wrt1.Write(ConnectionIdentifier)

        Wrt1.Write(TypeID2)
        Wrt1.Write(Length2)
        Wrt1.Write(SequenceCount)
        Wrt1.Write(SendData)

        Wrt1.Flush()

        Return Mem1.ToArray

    End Function

    Public Function Read(Buffer() As Byte) As Byte()

        Dim Mem1 As New System.IO.MemoryStream(Buffer)
        Dim Rdr1 As New System.IO.BinaryReader(Mem1)

        CommandR = Rdr1.ReadInt16
        LengthR = Rdr1.ReadInt16
        SessionHandleR = Rdr1.ReadInt32
        StatusR = Rdr1.ReadInt32
        SenderContextR = Rdr1.ReadInt64
        OptionsR = Rdr1.ReadInt32

        ' Read Data if not error

        If StatusR = 0 Then

            InterfaceHandleR = Rdr1.ReadInt32

            TimeOutR = Rdr1.ReadInt16
            ItemCountR = Rdr1.ReadInt16

            TypeID1R = Rdr1.ReadInt16
            Length1R = Rdr1.ReadInt16
            ConnectionIdentifierR = Rdr1.ReadInt32

            TypeID2R = Rdr1.ReadInt16
            Length2R = Rdr1.ReadInt16
            SequenceCountR = Rdr1.ReadInt16

            Dim RetLen As Integer = Mem1.Length - Mem1.Position
            If (RetLen <> (Length2R - 2)) Then Return Nothing

            Dim RetBuf(RetLen - 1) As Byte
            Array.Copy(Mem1.ToArray, Mem1.Position, RetBuf, 0, RetLen)

            Return RetBuf

        Else
            Return Nothing

        End If




    End Function



End Class
