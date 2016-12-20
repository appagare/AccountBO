Imports ASi.DataAccess.SqlHelper
Imports ASi.AccountBO.Constants

Friend NotInheritable Class AccountTypes
    Private _ConnectString As String = ""
    Private _AccountID As Integer = 0

    'could consider changing these to Datasets so they are serializable and then changing the Friend Properties to support both tables and datasets
    'need to consider Account class which maintains the entities as datatables (but, as long as DataTables are supported in addition to Datasets, it shouldn't matter).

    Private _AccountTypes As New DataTable
    Private _AddressTypes As New DataTable
    Private _PhoneTypes As New DataTable
    Private _PersonTypes As New DataTable
    Private _EmailTypes As New DataTable
    Private _RoleTypes As New DataTable
    Private _ServiceTypes As New DataTable
    Private _OrderStatusTypes As New DataTable
    Private _PackingListStatusTypes As New DataTable
    Private _POStatusTypes As New DataTable

    'Private _Constants As Constants
    
    Friend Property AccountID() As Integer
        Get
            Return _AccountID
        End Get
        Set(ByVal Value As Integer)
            _AccountID = Value
            Refresh()
        End Set
    End Property

    Friend ReadOnly Property AccountTypes() As DataTable
        Get
            Return _AccountTypes
        End Get
    End Property

    Friend ReadOnly Property AddressTypes() As DataTable
        Get
            Return _AddressTypes
        End Get
    End Property
    Friend ReadOnly Property PhoneTypes() As DataTable
        Get
            Return _PhoneTypes
        End Get
    End Property
    Friend ReadOnly Property PersonTypes() As DataTable
        Get
            Return _PersonTypes
        End Get
    End Property

    Friend ReadOnly Property EmailTypes() As DataTable
        Get
            Return _EmailTypes
        End Get
    End Property

    Friend ReadOnly Property RoleTypes() As DataTable
        Get
            Return _RoleTypes
        End Get
    End Property
    Friend ReadOnly Property ServiceTypes() As DataTable
        Get
            Return _ServiceTypes
        End Get
    End Property
    Friend ReadOnly Property OrderStatusTypes As DataTable
        Get
            Return _OrderStatusTypes
        End Get
    End Property
    Friend ReadOnly Property PackingListStatusTypes As DataTable
        Get
            Return _PackingListStatusTypes
        End Get
    End Property
    Friend ReadOnly Property POStatusTypes As DataTable
        Get
            Return _POStatusTypes
        End Get
    End Property



    Friend Sub Refresh()

        'refresh all of the lists
        If _AccountID > 0 Then
            'specific account
            _AccountTypes = ExecuteDataset(_ConnectString, _
                CommandType.StoredProcedure, _
                "AccountTypeCodeList", _
                New SqlClient.SqlParameter("@AccountID", _AccountID), _
                New SqlClient.SqlParameter("@Type", ACCOUNT_TYPE)).Tables(0)

            _EmailTypes = ExecuteDataset(_ConnectString, _
                CommandType.StoredProcedure, _
                "AccountTypeCodeList", _
                New SqlClient.SqlParameter("@AccountID", _AccountID), _
                New SqlClient.SqlParameter("@Type", EMAIL_TYPE)).Tables(0)

            _AddressTypes = ExecuteDataset(_ConnectString, _
                CommandType.StoredProcedure, _
                "AccountTypeCodeList", _
                New SqlClient.SqlParameter("@AccountID", _AccountID), _
                New SqlClient.SqlParameter("@Type", ADDRESS_TYPE)).Tables(0)

            _PhoneTypes = ExecuteDataset(_ConnectString, _
                CommandType.StoredProcedure, _
                "AccountTypeCodeList", _
                New SqlClient.SqlParameter("@AccountID", _AccountID), _
                New SqlClient.SqlParameter("@Type", PHONE_TYPE)).Tables(0)

            _PersonTypes = ExecuteDataset(_ConnectString, _
                CommandType.StoredProcedure, _
                "AccountTypeCodeList", _
                New SqlClient.SqlParameter("@AccountID", _AccountID), _
                New SqlClient.SqlParameter("@Type", PERSON_TYPE)).Tables(0)

            _RoleTypes = ExecuteDataset(_ConnectString, _
                CommandType.StoredProcedure, _
                "AccountTypeCodeList", _
                New SqlClient.SqlParameter("@AccountID", _AccountID), _
                New SqlClient.SqlParameter("@Type", ROLE_TYPE)).Tables(0)

            _ServiceTypes = ExecuteDataset(_ConnectString, _
                CommandType.StoredProcedure, _
                "AccountTypeCodeList", _
                New SqlClient.SqlParameter("@AccountID", _AccountID), _
                New SqlClient.SqlParameter("@Type", SERVICE_TYPE)).Tables(0)

            _OrderStatusTypes = ExecuteDataset(_ConnectString, _
                CommandType.StoredProcedure, _
                "AccountTypeCodeList", _
                New SqlClient.SqlParameter("@AccountID", _AccountID), _
                New SqlClient.SqlParameter("@Type", ORDER_STATUS_TYPE)).Tables(0)

            _PackingListStatusTypes = ExecuteDataset(_ConnectString, _
                CommandType.StoredProcedure, _
                "AccountTypeCodeList", _
                New SqlClient.SqlParameter("@AccountID", _AccountID), _
                New SqlClient.SqlParameter("@Type", PACKING_SLIP_STATUS_TYPE)).Tables(0)

            _POStatusTypes = ExecuteDataset(_ConnectString, _
                CommandType.StoredProcedure, _
                "AccountTypeCodeList", _
                New SqlClient.SqlParameter("@AccountID", _AccountID), _
                New SqlClient.SqlParameter("@Type", PO_STATUS_TYPE)).Tables(0)

        Else
            'defaults
            _AccountTypes = ExecuteDataset(_ConnectString, _
                CommandType.StoredProcedure, _
                "DefaultTypeCodeList", _
                New SqlClient.SqlParameter("@Type", ACCOUNT_TYPE)).Tables(0)

            _EmailTypes = ExecuteDataset(_ConnectString, _
                CommandType.StoredProcedure, _
                "DefaultTypeCodeList", _
                New SqlClient.SqlParameter("@Type", EMAIL_TYPE)).Tables(0)

            _AddressTypes = ExecuteDataset(_ConnectString, _
                CommandType.StoredProcedure, _
                "DefaultTypeCodeList", _
                New SqlClient.SqlParameter("@Type", ADDRESS_TYPE)).Tables(0)

            _PhoneTypes = ExecuteDataset(_ConnectString, _
                CommandType.StoredProcedure, _
                "DefaultTypeCodeList", _
                New SqlClient.SqlParameter("@Type", PHONE_TYPE)).Tables(0)

            _PersonTypes = ExecuteDataset(_ConnectString, _
                CommandType.StoredProcedure, _
                "DefaultTypeCodeList", _
                New SqlClient.SqlParameter("@Type", PERSON_TYPE)).Tables(0)

            _RoleTypes = ExecuteDataset(_ConnectString, _
                CommandType.StoredProcedure, _
                "DefaultTypeCodeList", _
                New SqlClient.SqlParameter("@Type", ROLE_TYPE)).Tables(0)

            'is this necessary or set it to something that will return empty _ServiceTypes?
            'updated 3/18/2010 - set to blank
            _ServiceTypes = ExecuteDataset(_ConnectString, _
                CommandType.StoredProcedure, _
                "DefaultTypeCodeList", _
                New SqlClient.SqlParameter("@Type", SERVICE_TYPE & BLANK)).Tables(0)

            'shouldn't happen
            _OrderStatusTypes = ExecuteDataset(_ConnectString, _
                CommandType.StoredProcedure, _
                "DefaultTypeCodeList", _
                New SqlClient.SqlParameter("@Type", ORDER_STATUS_TYPE)).Tables(0)

            _PackingListStatusTypes = ExecuteDataset(_ConnectString, _
                CommandType.StoredProcedure, _
                "DefaultTypeCodeList", _
                New SqlClient.SqlParameter("@Type", PACKING_SLIP_STATUS_TYPE)).Tables(0)

            _POStatusTypes = ExecuteDataset(_ConnectString, _
                CommandType.StoredProcedure, _
                "DefaultTypeCodeList", _
                New SqlClient.SqlParameter("@Type", PO_STATUS_TYPE)).Tables(0)

        End If
    End Sub

    Friend Sub New(ByVal ConnectionString As String, Optional ByVal AccountID As Integer = 0)
        _ConnectString = ConnectionString
        _AccountID = AccountID
        Refresh()
    End Sub

#Region "CRUD Functions"
    Friend Sub TypeDelete(ByVal Type As String, ByVal Code As String)
        'validation
        If Type = "" Then
            Throw New Exception("Type is required.")
        ElseIf UCase(Code) = BLANK Then
            Throw New Exception(BLANK & " is restricted and cannot be deleted.")
        End If
        'ok to delete
        ExecuteNonQuery(_ConnectString, _
                        CommandType.StoredProcedure, _
                        "AccountTypeCodeDelete", _
                        New SqlClient.SqlParameter("@AccountID", _AccountID), _
                        New SqlClient.SqlParameter("@Type", Type), _
                        New SqlClient.SqlParameter("@Code", Code))

        Refresh()

    End Sub
    Friend Sub AccountTypeDelete(ByVal Code As String)
        TypeDelete(ACCOUNT_TYPE, Code)
    End Sub
    Friend Sub AddressTypeDelete(ByVal Code As String)
        TypeDelete(ADDRESS_TYPE, Code)
    End Sub
    Friend Sub PersonTypeDelete(ByVal Code As String)
        TypeDelete(PERSON_TYPE, Code)
    End Sub
    Friend Sub PhoneTypeDelete(ByVal Code As String)
        TypeDelete(PHONE_TYPE, Code)
    End Sub
    Friend Sub EmailTypeDelete(ByVal Code As String)
        TypeDelete(EMAIL_TYPE, Code)
    End Sub
    Friend Sub RoleTypeDelete(ByVal Code As String)
        TypeDelete(ROLE_TYPE, Code)
    End Sub
    Friend Sub ServiceTypeDelete(ByVal Code As String)
        TypeDelete(SERVICE_TYPE, Code)
    End Sub

    Friend Sub TypeSet(ByVal Type As String, ByVal Code As String, _
        ByVal Description As String, ByVal SortOrder As Integer, _
        ByVal Comment As String)


        'validation
        If Type = "" Then
            Throw New Exception("Type is required.")
        ElseIf Code = "" Then
            Throw New Exception("Code is required.")
        End If

        ExecuteNonQuery(_ConnectString, _
                        CommandType.StoredProcedure, _
                        "AccountTypeCodeSet", _
                        New SqlClient.SqlParameter("@AccountID", _AccountID), _
                        New SqlClient.SqlParameter("@Type", Type), _
                        New SqlClient.SqlParameter("@Code", Code), _
                        New SqlClient.SqlParameter("@Description", Description), _
                        New SqlClient.SqlParameter("@SortOrder", SortOrder), _
                        New SqlClient.SqlParameter("@Comment", Comment))

        Refresh()

    End Sub
    Friend Sub AccountTypeSet(ByVal Code As String, _
        ByVal Description As String, ByVal SortOrder As Integer, _
        ByVal Comment As String)
        TypeSet(ACCOUNT_TYPE, Code, Description, SortOrder, Comment)
    End Sub
    Friend Sub AddressTypeSet(ByVal Code As String, _
        ByVal Description As String, ByVal SortOrder As Integer, _
        ByVal Comment As String)
        TypeSet(ADDRESS_TYPE, Code, Description, SortOrder, Comment)
    End Sub
    Friend Sub PersonTypeSet(ByVal Code As String, _
        ByVal Description As String, ByVal SortOrder As Integer, _
        ByVal Comment As String)
        TypeSet(PERSON_TYPE, Code, Description, SortOrder, Comment)
    End Sub
    Friend Sub PhoneTypeSet(ByVal Code As String, _
        ByVal Description As String, ByVal SortOrder As Integer, _
        ByVal Comment As String)
        TypeSet(PHONE_TYPE, Code, Description, SortOrder, Comment)
    End Sub
    Friend Sub EmailTypeSet(ByVal Code As String, _
        ByVal Description As String, ByVal SortOrder As Integer, _
        ByVal Comment As String)
        TypeSet(EMAIL_TYPE, Code, Description, SortOrder, Comment)
    End Sub
    Friend Sub RoleTypeSet(ByVal Code As String, _
        ByVal Description As String, ByVal SortOrder As Integer, _
        ByVal Comment As String)
        TypeSet(ROLE_TYPE, Code, Description, SortOrder, Comment)
    End Sub
    Friend Sub ServiceTypeSet(ByVal Code As String, _
        ByVal Description As String, ByVal SortOrder As Integer, _
        ByVal Comment As String)
        TypeSet(SERVICE_TYPE, Code, Description, SortOrder, Comment)
    End Sub

#End Region

End Class


