Public Module Constants

    Public Const IP_ADD As String = "10.244.105.199"
    'Public Const IP_ADD As String = "10.244.200.250"
    Public Const FilePath As String = "C:\Documents and Settings\jsala\Escritorio\CIP Capture\"


    Public Enum ENIP_Function_Codes As Int16
        NOP = &H0
        LIST_SERVICES = &H4
        LIST_IDENTITY = &H63
        LIST_INTERFACES = &H64
        REGISTER_SESSION = &H65
        UNREGISTER_SESSION = &H66
        SEND_RR_DATA = &H6F
        SEND_UNIT_DATA = &H70
        INDICATE_STATUS = &H72
        CANCEL = &H73
    End Enum

    Public Enum ENIP_Status_Codes As Int32
        SUCCESS = &H0
        INVALID_CMD = &H1
        NO_RESOURCES = &H2
        INCORRECT_DATA = &H3
        INVALID_SESSION = &H64
        INVALID_LENGTH = &H65
        UNSUPPORTED_PROT_REV = &H69
    End Enum

    Public Enum DataTypes

        BOOL = &HC1
        SINT = &HC2
        INT = &HC3
        DINT = &HC4

    End Enum

    Public Enum ItemTypes

        Null = 0
        ConnectedAddressItem = &HA1
        ConnectedDataItem = &HB1
        UnconnectedDataItem = &HB2
        SequencedAddressItem = &H8002

    End Enum

    Public Enum RA_Services

        MultiRequest = &H3
        GetAttributesList = &HA
        ExecutePCCC = &H4B
        CIP_ReadData = &H4C
        CIP_WriteData = &H4D
        ReadModifyWrite = &H4E
        CIP_ReadDataFragmented = &H52
        CIP_WriteDataFragmented = &H53

    End Enum

    Public Enum CIP_Services

        SC_GET_ATT_ALL = &H1
        SC_SET_ATT_ALL = &H2
        SC_GET_ATT_LIST = &H3
        SC_SET_ATT_LIST = &H4
        SC_RESET = &H5
        SC_START = &H6
        SC_STOP = &H7
        SC_CREATE = &H8
        SC_DELETE = &H9
        SC_MULT_SERV_PACK = &HA
        SC_APPLY_ATTRIBUTES = &HD
        SC_GET_ATT_SINGLE = &HE
        SC_SET_ATT_SINGLE = &H10
        SC_FIND_NEXT_OBJ_INST = &H11
        SC_RESTOR = &H15
        SC_SAVE = &H16
        SC_NO_OP = &H17
        SC_GET_MEMBER = &H18
        SC_SET_MEMBER = &H19
        SC_FWD_CLOSE = &H4E
        SC_UNCONECTED_SEND = &H52
        SC_FWD_OPEN = &H54

    End Enum



End Module
