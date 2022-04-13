Public Class ENIP_RegisterSessionClass1

    Private Command As ENIP_Function_Codes 'Int16
    Private Status As ENIP_Status_Codes 'Int32
    Public Length As Int16 = 4
    Public SessionHandle As Int32
    Public SenderContext As Int64 = 0 'Sequence number question - answer
    Public Options As Int32
    Public InterfaceHandle As Int32

    Public ProtocolVersion As Int16 = 1
    Public OptionFlags As Int16 = 0

    'Received Data

    Private CommandR As ENIP_Function_Codes 'Int16
    Public LengthR As Int16
    Public SessionHandleR As Int32
    Private StatusR As ENIP_Status_Codes 'Int32
    Public SenderContextR As Int64
    Public OptionsR As Int32

    Public ProtocolVersionR As Int16 = 1
    Public OptionFlagsR As Int16 = 0


    Public Sub New()

        Command = ENIP_Function_Codes.REGISTER_SESSION
        Status = ENIP_Status_Codes.SUCCESS
        Options = 0

    End Sub

    Public Function Write() As Byte()

        Dim Mem1 As New System.IO.MemoryStream
        Dim Wrt1 As New System.IO.BinaryWriter(Mem1)

        Wrt1.Write(Command)
        Wrt1.Write(Length)
        Wrt1.Write(SessionHandle)
        Wrt1.Write(Status)
        Wrt1.Write(SenderContext)
        Wrt1.Write(Options)

        Wrt1.Write(ProtocolVersion)
        Wrt1.Write(OptionFlags)

        Wrt1.Flush()

        Return Mem1.ToArray

    End Function

    Public Function Read(Buffer() As Byte) As Int32

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

            ProtocolVersion = Rdr1.ReadInt16
            OptionFlags = Rdr1.ReadInt16

        End If

        Return SessionHandleR

    End Function



End Class
