'todo: code this as it is being used.

#Region "Notes"
'Purpose - provide a convenient and fast method of programmatically accessing an accounts configuration settings.
'This can either contain all config settings or it can be inherited to contain a single group
'
'- list the account's configuration settings
'       - select an account's configuration setting
'       - update an account's configuration setting
'
#End Region

Imports ASi.DataAccess.SqlHelper
Imports ASi.AccountBO.Constants

Public Class ConfigSettings
    Inherits BaseConfigSettings
    'Private _ConnectString As String = ""
    'Private _AccountID As Integer = 0

    'convenient and fast retrieval of config settings
    'bygroup or by group and parameter
    '
    'Public Function GetValue(ByVal Group As Constants.ConfigurationGroups, ByVal Parameter As String) As String

    'End Function
    'Public Function GetValue(ByVal Group As String, ByVal Parameter As String) As String

    'End Function
    'Public Function ConfigurationList(Optional ByVal Group As String = "", _
    'Optional ByVal Parameter As String = "") As DataTable

    'If Group <> "" AndAlso Parameter <> "" Then
    '' accountid, group, and parameter
    'Return ExecuteDataset(_ConnectString, CommandType.StoredProcedure, _
    '"AccountConfigurationSelect", _
    'New SqlClient.SqlParameter("@AccountID", _AccountID), _
    'New SqlClient.SqlParameter("@Group", Group), _
    'New SqlClient.SqlParameter("@Parameter", Parameter)).Tables(0)

    'ElseIf Group <> "" Then
    '' accountid and group
    'Return ExecuteDataset(_ConnectString, CommandType.StoredProcedure, _
    '"AccountConfigSelect", _
    'New SqlClient.SqlParameter("@AccountID", _AccountID), _
    'New SqlClient.SqlParameter("@Group", Group)).Tables(0)
    'Else
    '' accountid only
    'Return ExecuteDataset(_ConnectString, CommandType.StoredProcedure, _
    '"AccountConfigurationSelect", New SqlClient.SqlParameter("@AccountID", _AccountID)).Tables(0)
    'End If

    'End Function
    'Public Sub ConfigurationDelete(ByVal Group As String, _
    'ByVal Parameter As String)

    ''deletes a record from the Accounts Config table
    'ExecuteNonQuery(_ConnectString, _
    'CommandType.StoredProcedure, _
    '"AccountConfigDelete", _
    'New SqlClient.SqlParameter("@AccountID", _AccountID), _
    'New SqlClient.SqlParameter("@Group", Group), _
    'New SqlClient.SqlParameter("@Parameter", Parameter))

    'End Sub

    'Public Sub ConfigurationSet(ByVal Group As String, _
    'ByVal Parameter As String, _
    'ByVal Value As String, _
    'Optional ByVal SortOrder As Byte = 0, _
    'Optional ByVal Comment As String = "")

    ''inserts or updates a record into the Accounts Config table
    'ExecuteNonQuery(_ConnectString, _
    'CommandType.StoredProcedure, _
    '"AccountConfigSet", _
    'New SqlClient.SqlParameter("@AccountID", _AccountID), _
    'New SqlClient.SqlParameter("@Group", Group), _
    'New SqlClient.SqlParameter("@Parameter", Parameter), _
    'New SqlClient.SqlParameter("@Value", Value), _
    'New SqlClient.SqlParameter("@SortOrder", SortOrder), _
    'New SqlClient.SqlParameter("@Comment", Comment))

    'End Sub

    Public Sub New(ByVal ConnectionString As String, _
                   ByVal AccountID As Integer, _
                   ByVal ParentID As Integer, _
                   ByVal Process As String)

        MyBase.New(ConnectionString, AccountID, ParentID, Process)

        '_ConnectString = ConnectionString
        '_AccountID = AccountID

    End Sub
End Class
