#Region "Object Notes"
'Q: what is the purpose of this the Person object
'A: provides the programmatic access to all aspects of a person
'   -! select the person's record
'   -! update the person record
'   -! manage the persons's address records
'       -! list the person's address records
'       -! select an individual person address record
'       -! add a new person address record
'       -! delete a person address record
'       -! update a person address record
'   -! manage the person's phone records
'       -! list the person's phone records
'       -! select an individual person phone record
'       -! add a new person phone record
'       -! delete a person phone record
'       -! update a person phone record
'   -! manage the person's email records
'       -! list the persons's email records
'       -! select an individual person email record
'       -! add a new person email record
'       -! delete a person email record
'       -! update a person email record
'
'- other considerations:
'   - user object
'   - user object contains roles
' Q: does some security need to be applied to make sure whomever calls this person has authority to access User object?

#End Region
Imports ASi.DataAccess.SqlHelper
Imports ASi.AccountBO.Constants
Imports ASi.LogEvent.LogEvent
Public NotInheritable Class Person

    Private Const APP_NAME As String = "ASi.AccountBO"
    Private Const PROCESS As String = "User"
    Private _LogEventBO As ASi.LogEvent.LogEvent

    Private _ConnectString As String = ""
    Private _AccountID As Integer = 0
    Private _PersonID As Integer = 0
    Private _Types As AccountTypes
    Private _Validation As New Validation

    Private _PersonDS As New DataSet 'cached dataset
    Private _AddressList As New DataTable
    Private _PhoneList As New DataTable
    Private _EmailList As New DataTable
    Private _PersonRow As DataRow
    'Private _UserRoles As New DataTable 'optional roles for informational purposes

#Region "Public Person"

#Region "Main Functions"
    'public methods and properties
    Public ReadOnly Property AccountID() As Integer
        Get
            Return _AccountID
        End Get
    End Property
    Public ReadOnly Property PersonID() As Integer
        Get
            Return _PersonID
        End Get
    End Property
    Public ReadOnly Property PersonDataSet() As DataSet
        Get
            Return _PersonDS
        End Get
    End Property
    Public ReadOnly Property PersonDataRow() As DataRow
        Get
            Return _PersonRow
        End Get
    End Property
    Public ReadOnly Property ConnectString() As String
        Get
            Return _ConnectString
        End Get
    End Property
    Public ReadOnly Property AddressList() As DataTable
        Get
            Return _AddressList
        End Get
    End Property
    Public ReadOnly Property EmailList() As DataTable
        Get
            Return _EmailList
        End Get
    End Property
    Public ReadOnly Property PhoneList() As DataTable
        Get
            Return _PhoneList
        End Get
    End Property
    Public Sub PersonUpdate(ByVal FirstName As String, _
       ByVal LastName As String, _
       Optional ByVal TypeCode As String = Constants.BLANK, _
       Optional ByVal Prefix As String = "", _
       Optional ByVal MI As String = "", _
       Optional ByVal Suffix As String = "")
        'update this person's base record

        Dim DebugString As String = "Person.Person.Update:"

        'validation - note validation values passed ByRef
        Dim ValidationString As String = _Validation.ValidPersonParameters(FirstName, LastName, TypeCode, Prefix, MI, Suffix)
        If ValidationString <> "" Then
            Throw New Exception(DebugString & ValidationString)
        End If

        'ok to proceed
        'add w/ inactivedate
        ExecuteNonQuery(_ConnectString, CommandType.StoredProcedure, _
            "PersonUpdate", _
            New SqlClient.SqlParameter("@AccountID", _AccountID), _
            New SqlClient.SqlParameter("@PersonID", _PersonID), _
            New SqlClient.SqlParameter("@TypeCode", TypeCode), _
            New SqlClient.SqlParameter("@Prefix", Prefix), _
            New SqlClient.SqlParameter("@FirstName", FirstName), _
            New SqlClient.SqlParameter("@MI", MI), _
            New SqlClient.SqlParameter("@LastName", LastName), _
            New SqlClient.SqlParameter("@Suffix", Suffix))

        'refresh the dataset
        _PersonDS = _GetPersonDataset(_PersonID, _AccountID)

    End Sub
    Public Function IsUser() As Boolean
        'returns whether or not this Person is a user
        If _PersonDS.Tables(PersonTableOrdinals.Users).Rows.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function
    Public Function UserID() As Integer
        'returns whether or not this Person is a user
        If IsUser() = True Then
            Return CType(_PersonDS.Tables(PersonTableOrdinals.Users).Rows(0)("UserID"), Integer)
        Else
            Return 0
        End If
    End Function
    Public Function GetUserObject() As ASi.AccountBO.User
        'returns an instances of the User class 
        'consider commenting out this function if not needed
        If IsUser() = True Then
            Return New ASi.AccountBO.User(_ConnectString, _AccountID, CType(_PersonDS.Tables(PersonTableOrdinals.Users).Rows(0)("UserID"), Integer))
        Else
            Return Nothing
        End If
    End Function
    Public Sub Refresh()

        'refresh all of the types
        _Types.Refresh()

        'refresh this persons's dataset and address, phone, email, etc. tables
        _PersonDS = _GetPersonDataset(_PersonID, _AccountID)
        _RefreshList(PersonTableOrdinals.Base, False)
        _RefreshList(PersonTableOrdinals.Address, False)
        _RefreshList(PersonTableOrdinals.Phone, False)
        _RefreshList(PersonTableOrdinals.Email, False)
        _RefreshList(PersonTableOrdinals.UserRoles, False)

    End Sub
    Public Sub New(ByVal ConnectionString As String, _
        ByVal AccountID As Integer, _
        ByVal PersonID As Integer)

        _ConnectString = ConnectionString
        _AccountID = AccountID
        _PersonID = PersonID
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
#End Region

#Region "Address Functions"
    Public Function AddressAdd(ByVal Address1 As String, _
       ByVal Address2 As String, _
       ByVal City As String, _
       ByVal StateCode As String, _
       ByVal PostalCode As String, _
       Optional ByVal TypeCode As String = Constants.BLANK, _
       Optional ByVal CountryCode As String = "") As Integer
        'adds an address to this Person, updates the AddressList, and returns the address id

        Dim DebugString As String = "Person.Address.Add:"

        'validation - note validation values passed ByRef
        Dim ValidationString As String = _Validation.ValidAddressParameters(Address1, Address2, City, StateCode, PostalCode, TypeCode, CountryCode)
        If ValidationString <> "" Then
            Throw New Exception(DebugString & ValidationString)
        End If

        'add
        Dim prm As New SqlClient.SqlParameter
        prm.Direction = ParameterDirection.ReturnValue
        Dim ID As Integer = ExecuteScalar(_ConnectString, CommandType.StoredProcedure, _
            "PersonAddressInsert", _
            New SqlClient.SqlParameter("@AccountID", _AccountID), _
            New SqlClient.SqlParameter("@PersonID", _PersonID), _
            New SqlClient.SqlParameter("@TypeCode", TypeCode), _
            New SqlClient.SqlParameter("@Address1", Address1), _
            New SqlClient.SqlParameter("@Address2", Address2), _
            New SqlClient.SqlParameter("@City", City), _
            New SqlClient.SqlParameter("@StateCode", StateCode), _
            New SqlClient.SqlParameter("@PostalCode", PostalCode), _
            New SqlClient.SqlParameter("@CountryCode", CountryCode), _
            prm)

        'make sure we find a returnvalue or ID from scalar
        If ID = 0 Then
            ID = CType(prm.Value, Integer)
        End If

        'update the dataset and address list
        _RefreshList(PersonTableOrdinals.Address, True)
        'return the ID
        Return ID
    End Function
    Public Sub AddressDelete(ByVal AddressID As Integer)
        'deletes an address from this person and updates the address list
        ExecuteNonQuery(_ConnectString, CommandType.StoredProcedure, _
            "PersonAddressDelete", _
            New SqlClient.SqlParameter("@AccountID", _AccountID), _
            New SqlClient.SqlParameter("@PersonID", _PersonID), _
            New SqlClient.SqlParameter("@AddressID", AddressID))
        'refresh the dataset and address list
        _RefreshList(PersonTableOrdinals.Address, True)
    End Sub
    Public Sub AddressUpdate(ByVal AddressID As Integer, _
        ByVal Address1 As String, _
        ByVal Address2 As String, _
        ByVal City As String, _
        ByVal StateCode As String, _
        ByVal PostalCode As String, _
        Optional ByVal TypeCode As String = Constants.BLANK, _
        Optional ByVal CountryCode As String = "")
        'updates an address record for this person and updates the address list

        Dim DebugString As String = "Person.Address.Update:"

        'validation - note validation values passed ByRef
        Dim ValidationString As String = _Validation.ValidAddressParameters(Address1, Address2, City, StateCode, PostalCode, TypeCode, CountryCode)
        If ValidationString <> "" Then
            Throw New Exception(DebugString & ValidationString)
        End If
        'updates
        ExecuteNonQuery(_ConnectString, CommandType.StoredProcedure, _
            "PersonAddressUpdate", _
            New SqlClient.SqlParameter("@AccountID", _AccountID), _
            New SqlClient.SqlParameter("@PersonID", _PersonID), _
            New SqlClient.SqlParameter("@AddressID", AddressID), _
            New SqlClient.SqlParameter("@TypeCode", TypeCode), _
            New SqlClient.SqlParameter("@Address1", Address1), _
            New SqlClient.SqlParameter("@Address2", Address2), _
            New SqlClient.SqlParameter("@City", City), _
            New SqlClient.SqlParameter("@StateCode", StateCode), _
            New SqlClient.SqlParameter("@PostalCode", PostalCode), _
            New SqlClient.SqlParameter("@CountryCode", CountryCode))
        'refresh the dataset and address list
        _RefreshList(PersonTableOrdinals.Address, True)
    End Sub
    Public Function AddressSelect(ByVal AddressID As Integer) As DataRow
        If _AddressList.Select("AddressID=" & AddressID.ToString).Length > 0 Then
            Return _AddressList.Select("AddressID=" & AddressID.ToString)(0)
        Else
            Return Nothing
        End If
    End Function
    Public Function AddressSearch(ByVal SearchExpression As String) As DataRow()
        If _AddressList.Select(SearchExpression).Length > 0 Then
            Return _AddressList.Select(SearchExpression)
        Else
            Return Nothing
        End If
    End Function
#End Region

#Region "Email Functions"
    Public Function EmailAdd(ByVal EmailName As String, _
            ByVal EmailAddress As String, _
            Optional ByVal TypeCode As String = Constants.BLANK) As Integer
        'adds an email to this person, updates the EmailList, and returns the EmailID
        Dim DebugString As String = "Person.Email.Add:"

        'validation - note validation values passed ByRef
        Dim ValidationString As String = _Validation.ValidEmailParameters(EmailName, EmailAddress, TypeCode)
        If ValidationString <> "" Then
            Throw New Exception(DebugString & ValidationString)
        End If

        'return the ID
        Dim prm As New SqlClient.SqlParameter
        prm.Direction = ParameterDirection.ReturnValue
        Dim ID As Integer = ExecuteScalar(_ConnectString, CommandType.StoredProcedure, _
                    "PersonEmailInsert", _
                    New SqlClient.SqlParameter("@AccountID", _AccountID), _
                    New SqlClient.SqlParameter("@PersonID", _PersonID), _
                    New SqlClient.SqlParameter("@TypeCode", TypeCode), _
                    New SqlClient.SqlParameter("@EmailName", EmailName), _
                    New SqlClient.SqlParameter("@EmailAddress", EmailAddress), _
                    prm)

        'make sure we find a returnvalue or ID from scalar
        If ID = 0 Then
            ID = CType(prm.Value, Integer)
        End If

        'update the dataset and list
        _RefreshList(PersonTableOrdinals.Email, True)
        'return the ID
        Return ID
    End Function
    Public Sub EmailDelete(ByVal EmailID As Integer)
        'deletes an email from the person and updates the email list
        ExecuteNonQuery(_ConnectString, CommandType.StoredProcedure, _
           "PersonEmailDelete", _
           New SqlClient.SqlParameter("@AccountID", _AccountID), _
           New SqlClient.SqlParameter("@PersonID", _PersonID), _
           New SqlClient.SqlParameter("@EmailID", EmailID))
        'update the dataset and email list
        _RefreshList(PersonTableOrdinals.Email, True)
    End Sub
    Public Sub EmailLink(ByVal EmailID As Integer)
        'links an account email to the person and updates the email list
        ExecuteNonQuery(_ConnectString, CommandType.StoredProcedure, _
           "PersonEmailLink", _
           New SqlClient.SqlParameter("@AccountID", _AccountID), _
           New SqlClient.SqlParameter("@PersonID", _PersonID), _
           New SqlClient.SqlParameter("@EmailID", EmailID))
        'update the dataset and email list
        _RefreshList(PersonTableOrdinals.Email, True)
    End Sub
    Public Sub EmailUpdate(ByVal EmailID As Integer, _
          ByVal EmailName As String, _
          ByVal EmailAddress As String, _
          Optional ByVal TypeCode As String = Constants.BLANK)
        'updates an email for this person, updates the EmailList
        Dim DebugString As String = "Person.Email.Update:"

        'validation - note validation values passed ByRef
        Dim ValidationString As String = _Validation.ValidEmailParameters(EmailName, EmailAddress, TypeCode)
        If ValidationString <> "" Then
            Throw New Exception(DebugString & ValidationString)
        End If

        ExecuteNonQuery(_ConnectString, CommandType.StoredProcedure, _
                    "PersonEmailUpdate", _
                    New SqlClient.SqlParameter("@AccountID", _AccountID), _
                    New SqlClient.SqlParameter("@PersonID", _PersonID), _
                    New SqlClient.SqlParameter("@EmailID", EmailID), _
                    New SqlClient.SqlParameter("@TypeCode", TypeCode), _
                    New SqlClient.SqlParameter("@EmailName", EmailName), _
                    New SqlClient.SqlParameter("@EmailAddress", EmailAddress))
        'refresh the dataset and list
        _RefreshList(PersonTableOrdinals.Email, True)
    End Sub
    Public Function EmailSelect(ByVal EmailID As Integer) As DataRow
        If _EmailList.Select("EmailID=" & EmailID.ToString).Length > 0 Then
            Return _EmailList.Select("EmailID=" & EmailID.ToString)(0)
        Else
            Return Nothing
        End If
    End Function
    Public Function EmailSearch(ByVal SearchExpression As String) As DataRow()
        If _EmailList.Select(SearchExpression).Length > 0 Then
            Return _EmailList.Select(SearchExpression)
        Else
            Return Nothing
        End If
    End Function
#End Region

#Region "Phone Functions"
    Public Function PhoneAdd(ByVal PhoneNumber As String, _
               Optional ByVal Extension As String = "", _
               Optional ByVal TypeCode As String = Constants.BLANK) As Integer
        'adds a phone record to this person, updates the PhoneList, and returns the PhoneID
        Dim DebugString As String = "Person.Phone.Add:"

        'validation - note validation values passed ByRef
        Dim ValidationString As String = _Validation.ValidPhoneParameters(PhoneNumber, Extension, TypeCode)
        If ValidationString <> "" Then
            Throw New Exception(DebugString & ValidationString)
        End If

        'return the ID
        Dim prm As New SqlClient.SqlParameter
        prm.Direction = ParameterDirection.ReturnValue
        Dim ID As Integer = ExecuteScalar(_ConnectString, CommandType.StoredProcedure, _
                    "PersonPhoneInsert", _
                    New SqlClient.SqlParameter("@AccountID", _AccountID), _
                    New SqlClient.SqlParameter("@PersonID", _PersonID), _
                    New SqlClient.SqlParameter("@TypeCode", TypeCode), _
                    New SqlClient.SqlParameter("@PhoneNumber", PhoneNumber), _
                    New SqlClient.SqlParameter("@Ext", Extension), _
                    prm)

        'make sure we find a returnvalue or ID from scalar
        If ID = 0 Then
            ID = CType(prm.Value, Integer)
        End If

        'update the dataset and list
        _RefreshList(PersonTableOrdinals.Phone, True)
        'return the ID
        Return ID
    End Function
    Public Sub PhoneDelete(ByVal PhoneID As Integer)
        'deletes a phone from the person and updates the Phone list
        ExecuteNonQuery(_ConnectString, CommandType.StoredProcedure, _
            "PersonPhoneDelete", _
            New SqlClient.SqlParameter("@AccountID", _AccountID), _
            New SqlClient.SqlParameter("@PersonID", _PersonID), _
            New SqlClient.SqlParameter("@PhoneID", PhoneID))
        'update the dataset and phone list
        _RefreshList(PersonTableOrdinals.Phone, True)
    End Sub
    Public Sub PhoneLink(ByVal PhoneID As Integer)
        'links an account phone to the person and updates the phone list
        ExecuteNonQuery(_ConnectString, CommandType.StoredProcedure, _
           "PersonPhoneLink", _
           New SqlClient.SqlParameter("@AccountID", _AccountID), _
           New SqlClient.SqlParameter("@PersonID", _PersonID), _
           New SqlClient.SqlParameter("@PhoneID", PhoneID))
        'update the dataset and email list
        _RefreshList(PersonTableOrdinals.Phone, True)
    End Sub
    Public Sub PhoneUpdate(ByVal PhoneID As Integer, _
               ByVal PhoneNumber As String, _
               Optional ByVal Extension As String = "", _
               Optional ByVal TypeCode As String = Constants.BLANK)
        'updates a phone record for this person and updates the PhoneList
        Dim DebugString As String = "Person.Phone.Update:"

        'validation - note validation values passed ByRef
        Dim ValidationString As String = _Validation.ValidPhoneParameters(PhoneNumber, Extension, TypeCode)
        If ValidationString <> "" Then
            Throw New Exception(DebugString & ValidationString)
        End If

        ExecuteNonQuery(_ConnectString, CommandType.StoredProcedure, _
                    "PersonPhoneUpdate", _
                    New SqlClient.SqlParameter("@AccountID", _AccountID), _
                    New SqlClient.SqlParameter("@PersonID", _PersonID), _
                    New SqlClient.SqlParameter("@PhoneID", PhoneID), _
                    New SqlClient.SqlParameter("@TypeCode", TypeCode), _
                    New SqlClient.SqlParameter("@PhoneNumber", PhoneNumber), _
                    New SqlClient.SqlParameter("@Ext", Extension))

        'update the dataset and list
        _RefreshList(PersonTableOrdinals.Phone, True)
    End Sub
    Public Function PhoneSelect(ByVal PhoneID As Integer) As DataRow
        If _PhoneList.Select("PhoneID=" & PhoneID.ToString).Length > 0 Then
            Return _PhoneList.Select("PhoneID=" & PhoneID.ToString)(0)
        Else
            Return Nothing
        End If
    End Function
    Public Function PhoneSearch(ByVal SearchExpression As String) As DataRow()
        If _PhoneList.Select(SearchExpression).Length > 0 Then
            Return _PhoneList.Select(SearchExpression)
        Else
            Return Nothing
        End If
    End Function
#End Region
#End Region


#Region "Private"
    'private methods and properties
    Private Function _GetPersonDataset(ByVal PersonID As Integer, ByVal AccountID As Integer) As DataSet
        'return an account dataset
        Return ExecuteDataset(_ConnectString, CommandType.StoredProcedure, "PersonSelect", _
                              New SqlClient.SqlParameter("@AccountID", AccountID), _
                              New SqlClient.SqlParameter("@PersonID", PersonID))

    End Function

    Private Sub _RefreshList(ByVal TableOrdinal As PersonTableOrdinals, ByVal RefreshDataSet As Boolean)
        'refresh this object's tables and optionally the complete dataset 
        If RefreshDataSet = True Then
            _PersonDS = _GetPersonDataset(_PersonID, _AccountID)
        End If
        Select Case TableOrdinal
            Case PersonTableOrdinals.Address
                _AddressList = _PersonDS.Tables(PersonTableOrdinals.Address)
            Case PersonTableOrdinals.Phone
                _PhoneList = _PersonDS.Tables(PersonTableOrdinals.Phone)
            Case PersonTableOrdinals.Email
                _EmailList = _PersonDS.Tables(PersonTableOrdinals.Email)
                'Case PersonTableOrdinals.UserRoles
                '_UserRoles = _PersonDS.Tables(PersonTableOrdinals.UserRoles)
            Case PersonTableOrdinals.Base
                If _PersonDS.Tables(PersonTableOrdinals.Base).Rows.Count > 0 Then
                    _PersonRow = _PersonDS.Tables(PersonTableOrdinals.Base).Rows(0)
                Else
                    Throw New Exception("There is no person row.")
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
