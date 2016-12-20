#Region "Object Notes"
'Q: what is the purpose of this the User object
'A: provides the programmatic access to all aspects of a User
'   - this is the main object that should be called by the service to authenticate and authorize functionality
'   !- select the user record
'   !- update the User record
'   !- list the User's roles
'   !- grant roles to the user
'   !- revoke roles from the user
'   !- AccountBO adds/deletes users
'
'TODO: handle forced password updates (PCI compliance)
#End Region
Imports ASi.DataAccess.SqlHelper
Imports ASi.AccountBO.Constants
Imports ASi.LogEvent.LogEvent
Public Class User

    Private Const APP_NAME As String = "ASi.AccountBO"
    Private Const PROCESS As String = "User"
    Private _LogEventBO As ASi.LogEvent.LogEvent

    Private _ConnectString As String = ""
    Private _AccountID As Integer = 0
    Private _UserID As Integer = 0
    'Private _PersonID As Integer = 0
    Private _Validation As New Validation

    Private _UserDS As New DataSet 'cached dataset
    Private _Types As AccountTypes
    Private _UserRoles As New DataTable
    Private _UserRolesHash As New Hashtable 'faster performance
    Private _UserRow As DataRow
    Private _IsPerson As Boolean = False
    Private _Util As New ASi.UtilityHelper.Utilities

    'Private _GUID As System.Guid


    'public methods and properties
    Public Function IsPerson() As Boolean
        Return _IsPerson
    End Function
    Public ReadOnly Property PersonID() As Integer
        Get
            If _IsPerson = True Then
                Return CType(_UserRow("PersonID"), Integer)
            Else
                Return 0
            End If
        End Get
    End Property
    Public ReadOnly Property AccountID() As Integer
        Get
            Return _AccountID
        End Get
    End Property
    Public ReadOnly Property UserID() As Integer
        Get
            Return _UserID
        End Get
    End Property
    Public ReadOnly Property UserDataSet() As DataSet
        Get
            Return _UserDS
        End Get
    End Property
    Public ReadOnly Property UserDataRow() As DataRow
        Get
            Return _UserRow
        End Get
    End Property
    Public ReadOnly Property ConnectString() As String
        Get
            Return _ConnectString
        End Get
    End Property
    Public ReadOnly Property RolesDataTable() As DataTable
        Get
            Return _UserRoles
        End Get
    End Property
    Public ReadOnly Property RolesHashTable() As Hashtable
        Get
            Return _UserRolesHash
        End Get
    End Property

    'Public Property SecurityToken() As String
    '    Get

    '    End Get
    '    Set(value As String)

    '    End Set
    'End Property


#Region "Security"
    'cache token
    Public Function IsAuthenticated() As Constants.TokenState
        'TODO:
        Return True
    End Function
    Public Function RoleIsAuthorized(ByVal Role As String) As Boolean
        Return _UserRolesHash.ContainsValue(Role.ToUpper)
    End Function
#End Region

    Public Sub UserUpdate(ByVal UserStatusID As UserStatus, _
                          ByVal UserName As String, _
                          ByVal EncryptedPassword As String, _
                          ByVal HashedPassword As String, _
                          ByVal Comment As String, _
                          ByVal Challenge As String, _
                          ByVal Response As String)

        'note - password should be encrypted at this point.
        Dim DebugString As String = "User.User.Update:"

        'validation - note validation values passed ByRef
        Dim ValidationString As String = _Validation.ValidUserUpdateParameters(UserName, EncryptedPassword)
        If ValidationString <> "" Then
            Throw New Exception(DebugString & ValidationString)
        End If

        'ok to proceed
        'add w/ inactivedate
        ExecuteNonQuery(_ConnectString, CommandType.StoredProcedure, _
            "UserUpdate", _
            New SqlClient.SqlParameter("@UserID", _UserID), _
            New SqlClient.SqlParameter("@UserStatusID", UserStatusID), _
            New SqlClient.SqlParameter("@UserName", UserName), _
            New SqlClient.SqlParameter("@EncryptedPassword", EncryptedPassword), _
            New SqlClient.SqlParameter("@HashedPassword", HashedPassword), _
            New SqlClient.SqlParameter("@Comment", Comment), _
            New SqlClient.SqlParameter("@Challenge", Challenge), _
            New SqlClient.SqlParameter("@Response", Response), _
            New SqlClient.SqlParameter("@PasswordExpireDays", 90))

        'refresh the dataset and base
        _RefreshList(UserTableOrdinals.Base, True)

    End Sub

    Public Function ValidPassword(ByVal PlainTextPassword As String) As Boolean
        Return _Util.StrongPassword(PlainTextPassword)
    End Function
    Public Function ValidUsername(ByVal Username As String, Optional ByVal ExistingUserID As Integer = 0) As Boolean

        'this is also in the Account.vb class
        Dim prm As New SqlClient.SqlParameter
        prm.Direction = ParameterDirection.ReturnValue
        Dim ID As Integer = ExecuteScalar(_ConnectString, CommandType.StoredProcedure, "UsernameExists", _
                                 New SqlClient.SqlParameter("@AccountID", _AccountID), _
                                 New SqlClient.SqlParameter("@Username", Username), _
                                 New SqlClient.SqlParameter("@UserID", ExistingUserID), _
                                 prm)

        'make sure we find a returnvalue or ID from scalar
        If ID = 0 And prm.Value <> 0 Then
            ID = CType(prm.Value, Integer)
        End If
        If ID = 0 Then
            'username record count - if zero, ok
            Return True
        Else
            'id > 0, username exists
            Return False
        End If

    End Function

    Public Sub LinkUserToPerson(ByVal PersonID As Integer)
        'links this user to a person
        ExecuteNonQuery(_ConnectString, CommandType.StoredProcedure, _
          "UserPersonLink", _
          New SqlClient.SqlParameter("@AccountID", _AccountID), _
          New SqlClient.SqlParameter("@UserID", _UserID), _
          New SqlClient.SqlParameter("@PersonID", PersonID))

        'update the dataset and roles
        _RefreshList(UserTableOrdinals.Base, True)

    End Sub
    Public Sub UnLinkUserFromPerson(ByVal PersonID As Integer)
        'unlinks this user from a person (this SP is also called in Account.UserDelete
        ExecuteNonQuery(_ConnectString, CommandType.StoredProcedure, _
          "UserPersonUnlink", _
          New SqlClient.SqlParameter("@AccountID", _AccountID), _
          New SqlClient.SqlParameter("@UserID", _UserID), _
          New SqlClient.SqlParameter("@PersonID", PersonID))

        'update the dataset and roles
        _RefreshList(UserTableOrdinals.Base, True)

    End Sub

    Public Sub GrantRole(ByVal Role As String)
        'make sure role exists in account's roles
        If _Types.RoleTypes.Select("Code='" & Role.ToUpper & "'").Length > 0 Then

            ExecuteNonQuery(_ConnectString, CommandType.StoredProcedure, _
          "UserRoleGrant", _
          New SqlClient.SqlParameter("@AccountID", _AccountID), _
          New SqlClient.SqlParameter("@UserID", _UserID), _
          New SqlClient.SqlParameter("@TypeCode", Role.ToUpper))

            'update the dataset and roles
            _RefreshList(UserTableOrdinals.UserRoles, False)
        Else
            Throw New Exception("[" & Role & "] is not a role that can be granted.")
        End If

    End Sub
    Public Sub RevokeRole(ByVal Role As String)

        ExecuteNonQuery(_ConnectString, CommandType.StoredProcedure, _
          "UserRoleRevoke", _
          New SqlClient.SqlParameter("@AccountID", _AccountID), _
          New SqlClient.SqlParameter("@UserID", _UserID), _
          New SqlClient.SqlParameter("@TypeCode", Role.ToUpper))

        'update the dataset and roles
        _RefreshList(UserTableOrdinals.UserRoles, False)
    End Sub

    Public Sub New(ByVal ConnectionString As String, _
        ByVal AccountID As Integer, _
        ByVal UserID As Integer)

        _ConnectString = ConnectionString
        _AccountID = AccountID
        _UserID = UserID
        _Types = New AccountTypes(_ConnectString, _AccountID)
        Refresh() 'refresh all objects in class

        Try
            _LogEventBO = New ASi.LogEvent.LogEvent
        Catch ex As Exception
            'consider throwing this error
            'if not, component may work w/out logging
            'if so, component will fail
        End Try

    End Sub
    Public Sub Refresh()

        'refresh all of the types
        _Types.Refresh()

        'refresh this users's dataset and roles, etc. tables
        _UserDS = _GetUserDataset(_UserID, _AccountID)

        _RefreshList(UserTableOrdinals.Base, False)
        _RefreshList(UserTableOrdinals.UserRoles, False)



    End Sub
#Region "Private"
    'private methods and properties
    Private Function _GetUserDataset(ByVal UserID As Integer, ByVal AccountID As Integer) As DataSet
        'return an account dataset
        Return ExecuteDataset(_ConnectString, CommandType.StoredProcedure, "UserSelect", _
                              New SqlClient.SqlParameter("@AccountID", AccountID), _
                              New SqlClient.SqlParameter("@UserID", UserID))

    End Function

    Private Sub _RefreshList(ByVal TableOrdinal As UserTableOrdinals, ByVal RefreshDataSet As Boolean)
        'refresh this object's tables and optionally the complete dataset 
        If RefreshDataSet = True Then
            _UserDS = _GetUserDataset(_UserID, _AccountID)
        End If
        Select Case TableOrdinal
            Case UserTableOrdinals.UserRoles
                _UserRoles = _UserDS.Tables(UserTableOrdinals.UserRoles)
                _UserRolesHash = _Util.DataSetToHashTable(_UserRoles, 1, 1, False, UtilityHelper.Utilities.TextCase.ToUpper)
            Case UserTableOrdinals.Base
                _IsPerson = False 'reset
                If _UserDS.Tables(UserTableOrdinals.Base).Rows.Count > 0 Then
                    _UserRow = _UserDS.Tables(UserTableOrdinals.Base).Rows(0)
                    If CType(_UserRow("PersonID"), Integer) > 0 Then
                        _IsPerson = True
                    End If
                Else
                    Throw New Exception("There is no user row.")
                End If
        End Select
    End Sub

    Private Sub _LogEvent(ByVal Src As String, ByVal Msg As String, ByVal Type As ASi.LogEvent.LogEvent.MessageType)
        Try
            _LogEventBO.LogEvent(APP_NAME, Src, Msg, Type, LogType.Queue)
        Catch ex As Exception
            _LogEventBO.LogEvent(APP_NAME, Src, Msg, Type, LogType.SystemEventLog)
        End Try
    End Sub
#End Region
End Class
