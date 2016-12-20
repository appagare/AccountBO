#Region "Object Notes"
'Terminology and Context:
'AccountID - this account's ID
'ParentID  - this account's ParentID. May be the same for root account
'Comments: 
'Account can update its own record but cannot change its own parent.
'Account can updates its direct child accounts by UI 
'ParentID cannot be changed programmatically. An account only needs to know it's parent ID for validation during the Select and Update SPs.
'When updating a child, the parent (this object) passes in it's account id  as the ParentID and the AccountID being modified is the child ID.
'
'An Account has configuration settings inherited from BaseConfigSettings
'Each service has configuration settings inherited from BaseConfigSettings
'
'Q: what is the purpose of this the AccountBO object
'A: provides the programmatic access to all aspects of an account
'   -! select the account's record
'   -! update the account record
'   -! manage the account's configuration settings - (indirectly by instantiating the ConfigSettings class)
'   -! manage the account's Types (AccountType, AddressType, EmailType, PhoneType, PersonType, and RoleType)
'       -! list the account's Types 
'       -! select an Type record 
'       -! set (add/update) a new account Type record 
'       -! delete an account Type record (but not BLANK) 
'   - manage the account's address records
'       -! list the account's address records
'       -! select a single account address record
'       -! add a new account address record
'       -! delete an account address record
'       -! update an account address record
'   - manage the account's phone records
'       -! list the account's phone records
'       -! select a single account phone record
'       -! add a new account phone record
'       -! delete an account phone record
'       -! update an account phone record
'   - manage the account's email records
'       -! list the account's email records
'       -! select a single account email record
'       -! add a new account email record
'       -! delete an account email record
'       -! update an account email record
'   - manage the account's person records
'       -! list the account's person records
'       -! select a single account person record (indirectly by instantiating an instance of the Person class)
'       -! add a new account person record
'       -! delete an account person record - todo: create SP
'   - manage the account's child accounts
'       -! add a new child record
'       -! update a child record (indirectly by instantiating another object to update the child record)
'       -! delete a child record - todo: create SP
'       -! list the account's child accounts
'       -! each child account should be managed by another instance of this object
'   - manage the account's user records
'       -! list the account's user records
'       -! select a single account user record (indirectly by instantiating an instance of the User class)
'       -! add a new account User record
'       -! delete an account User record - todo: create SP
'       -! roles managed via user object
'   - manage the account's service records
'       -! list the account's service records
'       -! select a single account service record (indirectly by instantiating an instance of the Service class)
'       -! add a new account Service record
'       -! delete an account Service record - todo: create SP

#End Region

Imports ASi.DataAccess.SqlHelper
Imports ASi.AccountBO.Constants
Imports ASi.LogEvent.LogEvent

Public NotInheritable Class Account

    Private Const APP_NAME As String = "ASi.AccountBO"
    Private Const PROCESS As String = "Account"
    Private _LogEventBO As ASi.LogEvent.LogEvent

    Private _ConnectString As String = ""
    Private _AccountID As Integer = 0
    Private _ParentID As Integer = 0

    Private _Types As AccountTypes
    Private _Validation As New Validation

    Private _AccountDS As New DataSet 'cached dataset
    Private _AddressList As New DataTable
    Private _PhoneList As New DataTable
    Private _PersonList As New DataTable
    Private _EmailList As New DataTable
    'Private _ChildList As New DataTable
    Private _UserList As New DataTable
    Private _NoteList As New DataTable
    Private _AccountRow As DataRow

    'todo: service class 
    '- type and config

#Region "Public Account"
    
#Region "Main Functions"
    'public methods and properties
    Public ReadOnly Property AccountID() As Integer
        Get
            Return _AccountID
        End Get
    End Property
    Public ReadOnly Property ParentID() As Integer
        Get
            Return _ParentID
        End Get
    End Property
    Public ReadOnly Property AccountDataSet() As DataSet
        Get
            Return _AccountDS
        End Get
    End Property
    Public ReadOnly Property AccountDataRow() As DataRow
        Get
            Return _AccountRow
        End Get
    End Property

    Public ReadOnly Property ShippingAddressRow() As DataRow
        Get
            'return a phone row (used by shipping labels ShippingUtility)
            'try shipping but if that doesn't exist, try _BLANK or whatever you can find
            If _AddressList.Select("TypeCode='SHIPPING'").Length > 0 Then
                Return _AddressList.Select("TypeCode='SHIPPING'")(0)
            ElseIf _AddressList.Select("TypeCode='_BLANK'").Length > 0 Then
                Return _AddressList.Select("TypeCode='_BLANK'")(0)
            Else
                If _AddressList.Rows.Count > 0 Then
                    'return whatever you can find
                    Return _AddressList.Rows(0)
                Else
                    'shouldn't happen
                    Return Nothing
                End If
            End If
        End Get
    End Property
    Public ReadOnly Property PhoneRow() As DataRow
        Get
            'return a phone row (used by shipping labels ShippingUtility)
            Return _PhoneList.Select("TypeCode='_BLANK'")(0)
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

    'Public ReadOnly Property ChildAccountList() As DataTable
    '    Get
    '        'change to a function - do not want to cache this - it can be huge
    '        Return _ChildList
    '    End Get
    'End Property
    Public Function ChildAccountList(Optional ByVal TypeCode As String = "", Optional ByVal AccountNameLike As String = "") As DataSet
        Return ExecuteDataset(_ConnectString, CommandType.StoredProcedure, "AccountChildList", _
                              New SqlClient.SqlParameter("@ParentID", AccountID), _
                              New SqlClient.SqlParameter("@TypeCode", TypeCode), _
                              New SqlClient.SqlParameter("@AccountNameLike", AccountNameLike))
    End Function

    Public Function DefaultStateList() As DataSet
        Return ExecuteDataset(_ConnectString, CommandType.StoredProcedure, "DefaultStateList")
    End Function
    Public Function DefaultCountryList() As DataSet
        Return ExecuteDataset(_ConnectString, CommandType.StoredProcedure, "DefaultCountryList")
    End Function
    Public Function CountryNameFromCode(ByVal CountryCode As String) As String
        Try
            Dim dt As New DataTable
            dt = DefaultCountryList.Tables(0)
            dt.DefaultView.RowFilter = "CountryCode='" & CountryCode & "'"
            If dt.DefaultView.Count = 1 Then
                Return dt.DefaultView(0)("Country")
            End If
            Return CountryCode
        Catch ex As Exception
            Return CountryCode
        End Try
    End Function
    Public Function CountryCodeFromName(ByVal CountryName As String) As String
        Try
            Dim dt As New DataTable
            dt = DefaultCountryList.Tables(0)
            dt.DefaultView.RowFilter = "Country='" & CountryName & "'"
            If dt.DefaultView.Count = 1 Then
                Return dt.DefaultView(0)("CountryCode")
            End If
            Return CountryName
        Catch ex As Exception
            Return CountryName
        End Try
    End Function
    Public Function StateNameFromCode(ByVal StateCode As String) As String
        Try
            Dim dt As New DataTable
            dt = DefaultStateList.Tables(0)
            dt.DefaultView.RowFilter = "StateCode='" & StateCode & "'"
            If dt.DefaultView.Count = 1 Then
                Return dt.DefaultView(0)("StateDescription")
            End If
            Return StateCode
        Catch ex As Exception
            Return StateCode
        End Try
    End Function
    Public Function StateCodeFromName(ByVal StateName As String) As String
        Try
            Dim dt As New DataTable
            dt = DefaultStateList.Tables(0)
            dt.DefaultView.RowFilter = "StateDescription='" & StateName & "'"
            If dt.DefaultView.Count = 1 Then
                Return dt.DefaultView(0)("StateCode")
            End If
            Return StateName
        Catch ex As Exception
            Return StateName
        End Try
    End Function

    Public ReadOnly Property EmailList() As DataTable
        Get
            Return _EmailList
        End Get
    End Property
    Public ReadOnly Property PersonList() As DataTable
        Get
            Return _PersonList
        End Get
    End Property
    Public ReadOnly Property PhoneList() As DataTable
        Get
            Return _PhoneList
        End Get
    End Property
    Public ReadOnly Property UserList() As DataTable
        Get
            Return _UserList
        End Get
    End Property
    Public ReadOnly Property NoteList() As DataTable
        Get
            Return _NoteList
        End Get
    End Property
    Public Sub AccountUpdate(ByVal AccountName As String, _
        ByVal URL As String, _
        Optional ByVal AccountStatus As Constants.AccountStatus = Constants.AccountStatus.Active, _
        Optional ByVal TypeCode As String = Constants.BLANK, _
        Optional ByVal InActiveDate As String = "", _
        Optional ByVal CustomTypeCode As String = "")
        'update this account's base record

        'validation - note validation values passed ByRef
        Dim ValidationString As String = _Validation.ValidAccountParameters(AccountName, URL, TypeCode, InActiveDate)
        If ValidationString <> "" Then
            Throw New Exception(ValidationString)
        End If

        'ok to proceed
        If InActiveDate <> "" AndAlso IsDate(InActiveDate) Then
            'add w/ inactivedate
            ExecuteNonQuery(_ConnectString, CommandType.StoredProcedure, _
                "AccountUpdate", _
                New SqlClient.SqlParameter("@AccountID", _AccountID), _
                New SqlClient.SqlParameter("@ParentID", _ParentID), _
                New SqlClient.SqlParameter("@TypeCode", TypeCode), _
                New SqlClient.SqlParameter("@AccountStatusID", AccountStatus), _
                New SqlClient.SqlParameter("@AccountName", AccountName), _
                New SqlClient.SqlParameter("@URL", URL), _
                New SqlClient.SqlParameter("@InActiveDate", CType(InActiveDate, Date)), _
                New SqlClient.SqlParameter("@CustomTypeCode", CustomTypeCode))

        Else
            'add w/out inactive date
            ExecuteNonQuery(_ConnectString, CommandType.StoredProcedure, _
                "AccountUpdate", _
                New SqlClient.SqlParameter("@AccountID", _AccountID), _
                New SqlClient.SqlParameter("@ParentID", _ParentID), _
                New SqlClient.SqlParameter("@TypeCode", TypeCode), _
                New SqlClient.SqlParameter("@AccountStatusID", AccountStatus), _
                New SqlClient.SqlParameter("@AccountName", AccountName), _
                New SqlClient.SqlParameter("@URL", URL), _
                New SqlClient.SqlParameter("@CustomTypeCode", CustomTypeCode))
        End If

        'refresh the dataset
        _AccountDS = _GetAccountDataset(_ParentID, _AccountID)

    End Sub
    'Public Function ConfigurationSettings() As ASi.AccountBO.ConfigSettings
    'Return New ASi.AccountBO.ConfigSettings(_ConnectString, _AccountID, _ParentID, PROCESS)
    'End Function
    Public Sub Refresh()

        'refresh all of the types
        _Types.Refresh()

        'refresh this account's dataset and address, phone, person, etc. tables
        _AccountDS = _GetAccountDataset(_ParentID, _AccountID)
        _RefreshList(AccountTableOrdinals.Base, False)
        _RefreshList(AccountTableOrdinals.Address, False)
        _RefreshList(AccountTableOrdinals.Phone, False)
        _RefreshList(AccountTableOrdinals.Email, False)
        _RefreshList(AccountTableOrdinals.Person, False)
        _RefreshList(AccountTableOrdinals.Users, False)
        _RefreshList(AccountTableOrdinals.Notes, False)
        '_RefreshList(AccountTableOrdinals.ChildAccounts, False)

    End Sub

    Public Sub New(ByVal ConnectionString As String, _
        ByVal ParentID As Integer, _
        ByVal AccountID As Integer)
        '_Process = Process
        _ConnectString = ConnectionString
        _AccountID = AccountID
        _ParentID = ParentID
        _Types = New AccountTypes(_ConnectString, _AccountID)
        Refresh()

        Try
            _LogEventBO = New ASi.LogEvent.LogEvent
        Catch ex As Exception
            'consider throwing this error
            'if not, component may work w/out logging
            'if so, component will fail
        End Try

    End Sub
#End Region
#Region "Notes"
    Public Function NoteAdd(ByVal NoteTitle As String, _
                            ByVal Visible As Boolean, _
                            ByVal Note As String, _
                            Optional ByVal UserID As Integer = 0, _
                            Optional ByVal TypeCode As String = Constants.ACCOUNT_TYPE, _
                            Optional ByVal Importance As Byte = 0) As Integer
        'adds a note to this account, updates the NoteList, and returns the note id

        Dim DebugString As String = "Account.Note.Add:"

        'validation - note validation values passed ByRef
        Dim ValidationString As String = _Validation.ValidNoteParameters(NoteTitle, Note)
        If ValidationString <> "" Then
            Throw New Exception(DebugString & ValidationString)
        End If

        'add
        Dim prm As New SqlClient.SqlParameter
        prm.Direction = ParameterDirection.ReturnValue
        Dim ID As Integer = ExecuteScalar(_ConnectString, CommandType.StoredProcedure, _
            "NoteInsert", _
            New SqlClient.SqlParameter("@AccountID", _AccountID), _
            New SqlClient.SqlParameter("@TypeCode", TypeCode), _
            New SqlClient.SqlParameter("@Importance", Importance), _
            New SqlClient.SqlParameter("@UserID", UserID), _
            New SqlClient.SqlParameter("@Visible", IIf(Visible = True, 1, 0)), _
            New SqlClient.SqlParameter("@NoteTitle", NoteTitle), _
            New SqlClient.SqlParameter("@Note", Note), _
            prm)

        'make sure we find a returnvalue or ID from scalar
        If ID = 0 Then
            ID = CType(prm.Value, Integer)
        End If

        'update the dataset and note list
        _RefreshList(AccountTableOrdinals.Notes, True)
        'return the ID
        Return ID
    End Function
    Public Sub NoteUpdate(ByVal NoteID As Integer, _
                               ByVal NoteTitle As String, _
                            ByVal Visible As Boolean, _
                            ByVal Note As String, _
                            Optional ByVal UserID As Integer = 0, _
                            Optional ByVal Importance As Byte = 0)
        'adds a note to this account, updates the NoteList, and returns the note id

        Dim DebugString As String = "Account.Note.Update:"

        'validation - note validation values passed ByRef
        Dim ValidationString As String = _Validation.ValidNoteParameters(NoteTitle, Note)
        If ValidationString <> "" Then
            Throw New Exception(DebugString & ValidationString)
        End If

        'update
        ExecuteNonQuery(_ConnectString, CommandType.StoredProcedure, _
                    "NoteUpdate", _
                    New SqlClient.SqlParameter("@NoteID", NoteID), _
                    New SqlClient.SqlParameter("@AccountID", _AccountID), _
            New SqlClient.SqlParameter("@Importance", Importance), _
            New SqlClient.SqlParameter("@UserID", UserID), _
            New SqlClient.SqlParameter("@Visible", IIf(Visible = True, 1, 0)), _
            New SqlClient.SqlParameter("@NoteTitle", NoteTitle), _
            New SqlClient.SqlParameter("@Note", Note))

        'update the dataset and note list
        _RefreshList(AccountTableOrdinals.Notes, True)

    End Sub
    Public Sub NoteDelete(ByVal NoteID As Integer)
        'marks a note as deleted and updates the notes list
        ExecuteNonQuery(_ConnectString, CommandType.StoredProcedure, _
            "NoteDelete", _
            New SqlClient.SqlParameter("@AccountID", _AccountID), _
            New SqlClient.SqlParameter("@NoteID", NoteID))
        'refresh the dataset and notes list
        _RefreshList(AccountTableOrdinals.Notes, True)
    End Sub
    Public Function NoteSelect(ByVal NoteID As Integer) As DataRow
        If _NoteList.Select("NoteID=" & NoteID.ToString).Length > 0 Then
            Return _NoteList.Select("NoteID=" & NoteID.ToString)(0)
        Else
            Return Nothing
        End If
    End Function
    Public Function NoteSearch(ByVal SearchExpression As String) As DataRow()
        If _NoteList.Select(SearchExpression).Length > 0 Then
            Return _NoteList.Select(SearchExpression)
        Else
            Return Nothing
        End If
    End Function
#End Region

#Region "Address Functions"
    Public Function AddressAdd(ByVal Address1 As String, _
       ByVal Address2 As String, _
       ByVal City As String, _
       ByVal StateCode As String, _
       ByVal PostalCode As String, _
       Optional ByVal TypeCode As String = Constants.BLANK, _
       Optional ByVal CountryCode As String = "") As Integer
        'adds an address to this account, updates the AddressList, and returns the address id

        Dim DebugString As String = "Account.Address.Add:"

        'validation - note validation values passed ByRef
        Dim ValidationString As String = _Validation.ValidAddressParameters(Address1, Address2, City, StateCode, PostalCode, TypeCode, CountryCode)
        If ValidationString <> "" Then
            Throw New Exception(DebugString & ValidationString)
        End If

        'add
        Dim prm As New SqlClient.SqlParameter
        prm.Direction = ParameterDirection.ReturnValue
        Dim ID As Integer = ExecuteScalar(_ConnectString, CommandType.StoredProcedure, _
            "AccountAddressInsert", _
            New SqlClient.SqlParameter("@AccountID", _AccountID), _
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
        _RefreshList(AccountTableOrdinals.Address, True)
        'return the ID
        Return ID
    End Function
    Public Sub AddressDelete(ByVal AddressID As Integer)
        'deletes an address from this account and updates the address list
        ExecuteNonQuery(_ConnectString, CommandType.StoredProcedure, _
            "AccountAddressDelete", _
            New SqlClient.SqlParameter("@AccountID", _AccountID), _
            New SqlClient.SqlParameter("@AddressID", AddressID))
        'refresh the dataset and address list
        _RefreshList(AccountTableOrdinals.Address, True)
    End Sub
    Public Sub AddressUpdate(ByVal AddressID As Integer, _
        ByVal Address1 As String, _
        ByVal Address2 As String, _
        ByVal City As String, _
        ByVal StateCode As String, _
        ByVal PostalCode As String, _
        Optional ByVal TypeCode As String = Constants.BLANK, _
        Optional ByVal CountryCode As String = "")
        'updates an address record for this account and updates the address list
        Dim DebugString As String = "Account.Address.Update:"

        'validation - note validation values passed ByRef
        Dim ValidationString As String = _Validation.ValidAddressParameters(Address1, Address2, City, StateCode, PostalCode, TypeCode, CountryCode)
        If ValidationString <> "" Then
            Throw New Exception(DebugString & ValidationString)
        End If
        'updates
        ExecuteNonQuery(_ConnectString, CommandType.StoredProcedure, _
            "AccountAddressUpdate", _
            New SqlClient.SqlParameter("@AccountID", _AccountID), _
            New SqlClient.SqlParameter("@AddressID", AddressID), _
            New SqlClient.SqlParameter("@TypeCode", TypeCode), _
            New SqlClient.SqlParameter("@Address1", Address1), _
            New SqlClient.SqlParameter("@Address2", Address2), _
            New SqlClient.SqlParameter("@City", City), _
            New SqlClient.SqlParameter("@StateCode", StateCode), _
            New SqlClient.SqlParameter("@PostalCode", PostalCode), _
            New SqlClient.SqlParameter("@CountryCode", CountryCode))
        'refresh the dataset and address list
        _RefreshList(AccountTableOrdinals.Address, True)
    End Sub
    Public Function AddressSelect(ByVal AddressID As Integer) As DataRow
        If _AddressList.Select("AddressID=" & AddressID.ToString).Length > 0 Then
            Return _AddressList.Select("AddressID=" & AddressID.ToString)(0)
        Else
            Return Nothing
        End If
    End Function
    Public Function AddressSearch(ByVal SearchExpression As String, Optional ByVal SortExpression As String = "") As DataRow()
        If _AddressList.Select(SearchExpression).Length > 0 Then
            Return _AddressList.Select(SearchExpression, SortExpression)
        Else
            Return Nothing
        End If
    End Function
#End Region

#Region "Email Functions"
    Public Function EmailAdd(ByVal EmailName As String, _
            ByVal EmailAddress As String, _
            Optional ByVal TypeCode As String = Constants.BLANK) As Integer
        'adds an email to this account, updates the EmailList, and returns the EmailID

        Dim DebugString As String = "Account.Email.Add:"

        'validation - note validation values passed ByRef
        Dim ValidationString As String = _Validation.ValidEmailParameters(EmailName, EmailAddress, TypeCode)
        If ValidationString <> "" Then
            Throw New Exception(DebugString & ValidationString)
        End If

        'return the ID
        Dim prm As New SqlClient.SqlParameter
        prm.Direction = ParameterDirection.ReturnValue
        Dim ID As Integer = ExecuteScalar(_ConnectString, CommandType.StoredProcedure, _
                    "AccountEmailInsert", _
                    New SqlClient.SqlParameter("@AccountID", _AccountID), _
                    New SqlClient.SqlParameter("@TypeCode", TypeCode), _
                    New SqlClient.SqlParameter("@EmailName", EmailName), _
                    New SqlClient.SqlParameter("@EmailAddress", EmailAddress), _
                    prm)

        'make sure we find a returnvalue or ID from scalar
        If ID = 0 Then
            ID = CType(prm.Value, Integer)
        End If

        'update the dataset and list
        _RefreshList(AccountTableOrdinals.Email, True)
        'return the ID
        Return ID
    End Function
    Public Sub EmailDelete(ByVal EmailID As Integer)
        'deletes an email from the account and updates the email list
        ExecuteNonQuery(_ConnectString, CommandType.StoredProcedure, _
           "AccountEmailDelete", _
           New SqlClient.SqlParameter("@AccountID", _AccountID), _
           New SqlClient.SqlParameter("@EmailID", EmailID))
        'update the dataset and email list
        _RefreshList(AccountTableOrdinals.Email, True)
    End Sub
    Public Sub EmailUpdate(ByVal EmailID As Integer, _
          ByVal EmailName As String, _
          ByVal EmailAddress As String, _
          Optional ByVal TypeCode As String = Constants.BLANK)
        'updates an email for this account, updates the EmailList

        Dim DebugString As String = "Account.Email.Update:"

        'validation - note validation values passed ByRef
        Dim ValidationString As String = _Validation.ValidEmailParameters(EmailName, EmailAddress, TypeCode)
        If ValidationString <> "" Then
            Throw New Exception(DebugString & ValidationString)
        End If

        ExecuteNonQuery(_ConnectString, CommandType.StoredProcedure, _
                    "AccountEmailUpdate", _
                    New SqlClient.SqlParameter("@AccountID", _AccountID), _
                    New SqlClient.SqlParameter("@EmailID", EmailID), _
                    New SqlClient.SqlParameter("@TypeCode", TypeCode), _
                    New SqlClient.SqlParameter("@EmailName", EmailName), _
                    New SqlClient.SqlParameter("@EmailAddress", EmailAddress))
        'refresh the dataset and list
        _RefreshList(AccountTableOrdinals.Email, True)
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
        'adds a phone record to this account, updates the PhoneList, and returns the PhoneID

        Dim DebugString As String = "Account.Phone.Add:"

        'validation - note validation values passed ByRef
        Dim ValidationString As String = _Validation.ValidPhoneParameters(PhoneNumber, Extension, TypeCode)
        If ValidationString <> "" Then
            Throw New Exception(DebugString & ValidationString)
        End If

        'return the ID
        Dim prm As New SqlClient.SqlParameter
        prm.Direction = ParameterDirection.ReturnValue
        Dim ID As Integer = ExecuteScalar(_ConnectString, CommandType.StoredProcedure, _
                    "AccountPhoneInsert", _
                    New SqlClient.SqlParameter("@AccountID", _AccountID), _
                    New SqlClient.SqlParameter("@TypeCode", TypeCode), _
                    New SqlClient.SqlParameter("@PhoneNumber", PhoneNumber), _
                    New SqlClient.SqlParameter("@Ext", Extension), _
                    prm)

        'make sure we find a returnvalue or ID from scalar
        If ID = 0 Then
            ID = CType(prm.Value, Integer)
        End If

        'update the dataset and list
        _RefreshList(AccountTableOrdinals.Phone, True)
        'return the ID
        Return ID
    End Function
    Public Sub PhoneDelete(ByVal PhoneID As Integer)
        'deletes a phone from the account and updates the Phone list
        ExecuteNonQuery(_ConnectString, CommandType.StoredProcedure, _
            "AccountPhoneDelete", _
            New SqlClient.SqlParameter("@AccountID", _AccountID), _
            New SqlClient.SqlParameter("@PhoneID", PhoneID))
        'update the dataset and phone list
        _RefreshList(AccountTableOrdinals.Phone, True)
    End Sub
    Public Sub PhoneUpdate(ByVal PhoneID As Integer, _
               ByVal PhoneNumber As String, _
               Optional ByVal Extension As String = "", _
               Optional ByVal TypeCode As String = Constants.BLANK)
        'updates a phone record for this account and updates the PhoneList
        Dim DebugString As String = "Account.Phone.Update:"

        'validation - note validation values passed ByRef
        Dim ValidationString As String = _Validation.ValidPhoneParameters(PhoneNumber, Extension, TypeCode)
        If ValidationString <> "" Then
            Throw New Exception(DebugString & ValidationString)
        End If

        ExecuteNonQuery(_ConnectString, CommandType.StoredProcedure, _
                    "AccountPhoneUpdate", _
                    New SqlClient.SqlParameter("@AccountID", _AccountID), _
                    New SqlClient.SqlParameter("@PhoneID", PhoneID), _
                    New SqlClient.SqlParameter("@TypeCode", TypeCode), _
                    New SqlClient.SqlParameter("@PhoneNumber", PhoneNumber), _
                    New SqlClient.SqlParameter("@Ext", Extension))

        'update the dataset and list
        _RefreshList(AccountTableOrdinals.Phone, True)
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
    Public Function PhoneFindPhoneIDByPhoneNumber(ByVal PhoneNumber As String) As Integer

        'slower but more reliable way to find an account PhoneID by the phone number
        PhoneNumber = Replace(Replace(Replace(Replace(Replace(PhoneNumber, " ", ""), "-", ""), ".", ""), "(", ""), "(", "")

        Dim ReturnValue As Integer = 0
        Dim r As DataRow
        For Each r In _PhoneList.Rows
            Dim TestNumber As String = Replace(Replace(Replace(Replace(Replace(r("PhoneNumber"), " ", ""), "-", ""), ".", ""), "(", ""), "(", "")
            If TestNumber = PhoneNumber Then
                ReturnValue = r(0) 'first column is PhoneID
                Exit For
            End If
        Next

        Return ReturnValue
    End Function
#End Region

#Region "Person Functions"
    Public Function PersonAdd(ByVal FirstName As String, _
       ByVal LastName As String, _
       Optional ByVal TypeCode As String = Constants.BLANK, _
       Optional ByVal Prefix As String = "", _
       Optional ByVal MI As String = "", _
       Optional ByVal Suffix As String = "") As Integer
        'adds a new person record to this account 
        Dim DebugString As String = "Account.Person.Add:"

        'validation - note validation values passed ByRef
        Dim ValidationString As String = _Validation.ValidPersonParameters(FirstName, LastName, TypeCode, Prefix, MI, Suffix)
        If ValidationString <> "" Then
            Throw New Exception(DebugString & ValidationString)
        End If

        'ok to proceed
        Dim prm As New SqlClient.SqlParameter
        prm.Direction = ParameterDirection.ReturnValue
        Dim ID As Integer = ExecuteScalar(_ConnectString, CommandType.StoredProcedure, _
                "PersonInsert", _
                New SqlClient.SqlParameter("@AccountID", _AccountID), _
                New SqlClient.SqlParameter("@TypeCode", TypeCode), _
                New SqlClient.SqlParameter("@Prefix", Prefix), _
                New SqlClient.SqlParameter("@FirstName", FirstName), _
                New SqlClient.SqlParameter("@MI", MI), _
                New SqlClient.SqlParameter("@LastName", LastName), _
                New SqlClient.SqlParameter("@Suffix", Suffix), prm)

        'make sure we find a returnvalue or ID from scalar
        If ID = 0 Then
            ID = CType(prm.Value, Integer)
        End If

        'update the dataset and list
        _RefreshList(AccountTableOrdinals.Person, True)
        'return the ID
        Return ID
    End Function
    Public Sub PersonDelete(ByVal PersonID As Integer)
        'deletes a person from this account
        
        If _UserList.Select("PersonID=" & PersonID.ToString).Length = 1 Then
            'if person is a user, user is also deleted
            Me.UserDelete(_UserList.Select("PersonID=" & PersonID.ToString)(0)("UserID"))
        End If

        'flags this person as deleted
        ExecuteNonQuery(_ConnectString, CommandType.StoredProcedure, _
           "PersonDelete", _
           New SqlClient.SqlParameter("@AccountID", _AccountID), _
           New SqlClient.SqlParameter("@PersonID", PersonID))

        'update the dataset and list
        _RefreshList(AccountTableOrdinals.Person, True)

    End Sub
    Public Sub PersonUpdate(ByVal PersonID As Integer, _
        ByVal FirstName As String, _
        ByVal LastName As String, _
        Optional ByVal TypeCode As String = Constants.BLANK, _
        Optional ByVal Prefix As String = "", _
        Optional ByVal MI As String = "", _
        Optional ByVal Suffix As String = "")
        'update this person's base record

        Dim DebugString As String = "Account.Person.Update:"

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
            New SqlClient.SqlParameter("@PersonID", PersonID), _
            New SqlClient.SqlParameter("@TypeCode", TypeCode), _
            New SqlClient.SqlParameter("@Prefix", Prefix), _
            New SqlClient.SqlParameter("@FirstName", FirstName), _
            New SqlClient.SqlParameter("@MI", MI), _
            New SqlClient.SqlParameter("@LastName", LastName), _
            New SqlClient.SqlParameter("@Suffix", Suffix))

        'update the dataset and list
        _RefreshList(AccountTableOrdinals.Person, True)

    End Sub

    Public Function PersonSelect(ByVal PersonID As Integer) As ASi.AccountBO.Person
        'returns an instances of the Person class 
        Return New ASi.AccountBO.Person(_ConnectString, _AccountID, PersonID)
    End Function

    'utility functions
    Public Function PersonEmailExists(ByVal PersonID As Integer, ByVal EmailAddress As String) As Integer

        'email/phone can be added/removed at the person level
        'if email/phone exists for another person, return their PersonID
        'if email/phone exists at the account level, just link it to the person
        'if email/phone doesn't exist, add it
        Dim ReturnValue As Integer = 0
        _PersonList.DefaultView.RowFilter = "EmailAddress='" & EmailAddress & "' and PersonID <> " & PersonID.ToString
        If _PersonList.DefaultView.Count > 0 Then
            ReturnValue = _PersonList.DefaultView(0)(0) 'first column is PersonID
        End If
        _PersonList.DefaultView.RowFilter = ""

        Return ReturnValue
    End Function

    Public Function PersonPhoneExists(ByVal PersonID As Integer, ByVal PhoneNumber As String) As Integer

        'email/phone can be added/removed at the person level
        'if email/phone exists for another person, return their PersonID
        'if email/phone exists at the account level, just link it to the person
        'if email/phone doesn't exist, add it
        'PhoneNumber = Replace(Replace(Replace(Replace(Replace(PhoneNumber, " ", ""), "-", ""), ".", ""), "(", ""), "(", "")

        'Dim ReturnValue As Integer = 0
        '_PersonList.DefaultView.RowFilter = "PersonID <> " & PersonID.ToString
        'If _PersonList.DefaultView.Count > 0 Then
        '    Dim i As Integer = 0
        '    For i = 0 To _PersonList.DefaultView.Count - 1
        '        Dim TestNumber As String = _PersonList.DefaultView(i)(_PersonList.Columns.Count - 1) 'AccountSelect SP should have Phone in 2nd to last column
        '        If TestNumber <> "" Then
        '            TestNumber = Replace(Replace(Replace(Replace(Replace(TestNumber, " ", ""), "-", ""), ".", ""), "(", ""), "(", "")
        '            If TestNumber = PhoneNumber Then
        '                ReturnValue = _PersonList.DefaultView(0)(0) 'first column is PersonID
        '                Exit For
        '            End If
        '        End If
        '    Next
        'End If
        '_PersonList.DefaultView.RowFilter = ""

        'slower but more reliable way to find an account PhoneID by the phone number
        PhoneNumber = Replace(Replace(Replace(Replace(Replace(PhoneNumber, " ", ""), "-", ""), ".", ""), "(", ""), "(", "")
        Dim ReturnValue As Integer = 0
        Dim r As DataRow
        For Each r In _PersonList.Rows
            Dim TestNumber As String = Replace(Replace(Replace(Replace(Replace(r("PhoneNumber"), " ", ""), "-", ""), ".", ""), "(", ""), "(", "")
            If TestNumber = PhoneNumber AndAlso r(0) <> PersonID Then
                ReturnValue = r(0) 'other person's PersonID is first column 
                Exit For
            End If
        Next
        Return ReturnValue

    End Function

#End Region

#Region "Child Account Functions"
    Public Function ChildAccountAdd(ByVal AccountName As String, _
       ByVal URL As String, _
       Optional ByVal AccountStatus As Constants.AccountStatus = Constants.AccountStatus.Active, _
       Optional ByVal TypeCode As String = Constants.BLANK, _
       Optional ByVal InActiveDate As String = "", _
       Optional ByVal CustomTypeCode As String = "") As Integer
        'adds a new child account to this account 
        Dim DebugString As String = "Account.ChildAccount.Add:"

        'validation - note validation values passed ByRef
        Dim ValidationString As String = _Validation.ValidAccountParameters(AccountName, URL, TypeCode, InActiveDate)
        If ValidationString <> "" Then
            Throw New Exception(DebugString & ValidationString)
        End If

        Dim ID As Integer = 0
        Dim prm As New SqlClient.SqlParameter
        prm.Direction = ParameterDirection.ReturnValue

        'ok to proceed
        If InActiveDate <> "" AndAlso IsDate(InActiveDate) Then
            'add w/ inactivedate
            ID = ExecuteScalar(_ConnectString, CommandType.StoredProcedure, _
                "AccountInsert", _
                New SqlClient.SqlParameter("@ParentID", _AccountID), _
                New SqlClient.SqlParameter("@TypeCode", TypeCode), _
                New SqlClient.SqlParameter("@AccountStatusID", AccountStatus), _
                New SqlClient.SqlParameter("@AccountName", AccountName), _
                New SqlClient.SqlParameter("@URL", URL), _
                New SqlClient.SqlParameter("@InActiveDate", CType(InActiveDate, Date)), _
                New SqlClient.SqlParameter("@CustomTypeCode", CustomTypeCode), _
                prm)

        Else
            'add w/out inactive date
            ID = ExecuteScalar(_ConnectString, CommandType.StoredProcedure, _
                "AccountInsert", _
                New SqlClient.SqlParameter("@ParentID", _AccountID), _
                New SqlClient.SqlParameter("@TypeCode", TypeCode), _
                New SqlClient.SqlParameter("@AccountStatusID", AccountStatus), _
                New SqlClient.SqlParameter("@AccountName", AccountName), _
                New SqlClient.SqlParameter("@URL", Trim(URL)), _
                New SqlClient.SqlParameter("@CustomTypeCode", CustomTypeCode), _
                prm)

        End If

        'update the dataset and list -updated 12/2/2012; removed ChildList from resultset; query separtely
        '_RefreshList(AccountTableOrdinals.ChildAccounts, True)

        If ID = 0 Then
            ID = CType(prm.Value, Integer)
        End If

        'return the ID
        Return ID
    End Function
    Public Sub ChildAccountDelete(ByVal ChildAccountID As Integer)
        'deletes a child account from this account
        'Todo: update SP - it's not done yet
        'SP will delete all data components associated with this account such as:
        ' NOTE - consider marking records "DELETED" and retain data for X period for recovery
        '- account types
        '- account options
        '- addresses
        '- email
        '- phone
        '- person
        ' - person address
        ' - person email
        ' - person phone
        ' - users
        ' - user roles
        ' - config settings
        ' - orders

        ExecuteNonQuery(_ConnectString, CommandType.StoredProcedure, _
           "AccountDelete", _
           New SqlClient.SqlParameter("@ParentID", _AccountID), _
           New SqlClient.SqlParameter("@AccountID", ChildAccountID))

        'update the dataset and child list -updated 12/2/2012; removed ChildList from resultset; query separtely
        '_RefreshList(AccountTableOrdinals.ChildAccounts, True)

    End Sub
    Public Function ChildAccountSelect(ByVal ChildAccountID As Integer) As ASi.AccountBO.Account
        'returns a new instances of this component containing the ChildAccount using this Account as the Parent
        Return New ASi.AccountBO.Account(_ConnectString, _AccountID, ChildAccountID)
    End Function
#End Region

#Region "User Functions"
    Public Function ValidUsername(ByVal Username As String, Optional ByVal ExistingUserID As Integer = 0) As Boolean
        'this is also in the User.vb class
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
    Public Function UserAdd(ByVal UserStatusID As UserStatus, _
                          ByVal UserName As String, _
                          ByVal EncryptedPassword As String, _
                          ByVal HashedPassword As String, _
                          ByVal Comment As String, _
                          ByVal Challenge As String, _
                          ByVal Response As String) As Integer
        'adds a new user record to this account 
        Dim DebugString As String = "Account.User.Add:"

        'validation - note validation values passed ByRef
        Dim ValidationString As String = _Validation.ValidUserUpdateParameters(UserName, EncryptedPassword)
        If ValidationString <> "" Then
            Throw New Exception(DebugString & ValidationString)
        End If

        'ok to proceed
        Dim prm As New SqlClient.SqlParameter
        prm.Direction = ParameterDirection.ReturnValue
        Dim ID As Integer = ExecuteScalar(_ConnectString, CommandType.StoredProcedure, _
                "UserInsert", _
                New SqlClient.SqlParameter("@AccountID", _AccountID), _
                New SqlClient.SqlParameter("@UserStatusID", UserStatusID), _
                New SqlClient.SqlParameter("@UserName", UserName), _
                New SqlClient.SqlParameter("@EncryptedPassword", EncryptedPassword), _
                New SqlClient.SqlParameter("@HashedPassword", HashedPassword), _
                New SqlClient.SqlParameter("@Comment", Comment), _
                New SqlClient.SqlParameter("@Challenge", Challenge), _
                New SqlClient.SqlParameter("@Response", Response), _
                New SqlClient.SqlParameter("@PasswordExpireDays", 90), _
                prm)

        'make sure we find a returnvalue or ID from scalar
        If ID = 0 Then
            ID = CType(prm.Value, Integer)
        End If

        'update the dataset and list
        _RefreshList(AccountTableOrdinals.Users, True)
        'return the ID
        Return ID
    End Function
    Public Sub UserDelete(ByVal UserID As Integer)
        'flags a user record as deleted. 
        Try
            'if User is a Person, also break the link to the person.
            'this allows the person to exist while no longer being a user
            Dim PersonID As Integer = CType(_UserList.Select("UserID=" & UserID.ToString)(0)("PersonID"), Integer)
            If PersonID > 0 Then
                'this code also exists in User class
                ExecuteNonQuery(_ConnectString, CommandType.StoredProcedure, _
                "UserPersonUnlink", _
                New SqlClient.SqlParameter("@AccountID", _AccountID), _
                New SqlClient.SqlParameter("@UserID", _UserList.Select("PersonID=" & PersonID.ToString)(0)("UserID")), _
                New SqlClient.SqlParameter("@PersonID", PersonID))
            End If
        Catch ex As Exception
        End Try

        'flag user as deleted
        ExecuteNonQuery(_ConnectString, CommandType.StoredProcedure, _
           "UserDelete", _
           New SqlClient.SqlParameter("@AccountID", _AccountID), _
           New SqlClient.SqlParameter("@UserID", UserID))

        'update the dataset and list
        _RefreshList(AccountTableOrdinals.Users, True)

    End Sub

    Public Function UserSelect(ByVal UserID As Integer) As ASi.AccountBO.User
        'returns an instances of the User class 
        Return New ASi.AccountBO.User(_ConnectString, _AccountID, UserID)
    End Function
#End Region

#End Region

#Region "Types"
    '*****************************************************
    ' AccountTypes 
    '*****************************************************
    Public ReadOnly Property AccountTypes() As DataTable
        Get
            'lists all Account Types
            Return _Types.AccountTypes
        End Get
    End Property
    Public Sub AccountTypeSet(ByVal TypeCode As String, _
       ByVal Description As String, _
       Optional ByVal SortOrder As Byte = 0, _
       Optional ByVal Comment As String = "")
        'adds or updates an Account Type record
        _Types.AccountTypeSet(TypeCode, Description, SortOrder, Comment)
    End Sub
    Public Function AccountTypeSelect(ByVal TypeCode As String) As DataRow
        'returns an Account Type row
        Return _TypeSelect(ACCOUNT_TYPE, TypeCode)
    End Function
    Public Sub AccountTypeDelete(ByVal TypeCode As String)
        'deletes an Account Type row
        _Types.AccountTypeDelete(TypeCode)
    End Sub


    '*****************************************************
    ' AddressTypes 
    '*****************************************************
    Public ReadOnly Property AddressTypes() As DataTable
        Get
            'lists all Address Types
            Return _Types.AddressTypes
        End Get
    End Property
    Public Sub AddressTypeSet(ByVal TypeCode As String, _
        ByVal Description As String, _
        Optional ByVal SortOrder As Byte = 0, _
        Optional ByVal Comment As String = "")
        'adds or updates an Address Type
        _Types.AddressTypeSet(TypeCode, Description, SortOrder, Comment)
    End Sub
    Public Function AddressTypeSelect(ByVal TypeCode As String) As DataRow
        'returns an Address Type row
        Return _TypeSelect(ADDRESS_TYPE, TypeCode)
    End Function
    Public Sub AddressTypeDelete(ByVal TypeCode As String)
        'deletes an Address Type row
        _Types.AddressTypeDelete(TypeCode)
    End Sub


    '*****************************************************
    ' EmailTypes 
    '*****************************************************
    Public ReadOnly Property EmailTypes() As DataTable
        Get
            'lists all Email Types
            Return _Types.EmailTypes
        End Get
    End Property
    Public Sub EmailTypeSet(ByVal TypeCode As String, _
       ByVal Description As String, _
       Optional ByVal SortOrder As Byte = 0, _
       Optional ByVal Comment As String = "")
        'adds or updates an Email Type
        _Types.EmailTypeSet(TypeCode, Description, SortOrder, Comment)
    End Sub
    Public Function EmailTypeSelect(ByVal TypeCode As String) As DataRow
        'returns an Email Type row
        Return _TypeSelect(EMAIL_TYPE, TypeCode)
    End Function
    Public Sub EmailTypeDelete(ByVal TypeCode As String)
        'deletes an Email Type row
        _Types.EmailTypeDelete(TypeCode)
    End Sub


    '*****************************************************
    ' PersonTypes 
    '*****************************************************
    Public ReadOnly Property PersonTypes() As DataTable
        Get
            'lists all Person Types
            Return _Types.PersonTypes
        End Get
    End Property
    Public Sub PersonTypeSet(ByVal TypeCode As String, _
       ByVal Description As String, _
       Optional ByVal SortOrder As Byte = 0, _
       Optional ByVal Comment As String = "")
        'adds or updates a Person Type
        _Types.PersonTypeSet(TypeCode, Description, SortOrder, Comment)
    End Sub
    Public Function PersonTypeSelect(ByVal TypeCode As String) As DataRow
        'returns a Person Type row
        Return _TypeSelect(PERSON_TYPE, TypeCode)
    End Function
    Public Sub PersonTypeDelete(ByVal TypeCode As String)
        'deletes a Person Type row
        _Types.PersonTypeDelete(TypeCode)
    End Sub


    '*****************************************************
    ' PhoneTypes 
    '*****************************************************
    Public ReadOnly Property PhoneTypes() As DataTable
        Get
            'lists all Phone Types
            Return _Types.PhoneTypes
        End Get
    End Property
    Public Sub PhoneTypeSet(ByVal TypeCode As String, _
        ByVal Description As String, _
        Optional ByVal SortOrder As Byte = 0, _
        Optional ByVal Comment As String = "")
        'adds or updates a Phone Type record
        _Types.PhoneTypeSet(TypeCode, Description, SortOrder, Comment)
    End Sub
    Public Function PhoneTypeSelect(ByVal TypeCode As String) As DataRow
        'returns a Phone Type record
        Return _TypeSelect(PHONE_TYPE, TypeCode)
    End Function
    Public Sub PhoneTypeDelete(ByVal TypeCode As String)
        'deletes a Phone Type record
        _Types.PhoneTypeDelete(TypeCode)
    End Sub


    '*****************************************************
    ' RoleTypes 
    '*****************************************************
    Public ReadOnly Property RoleTypes() As DataTable
        Get
            'lists all Role Types
            Return _Types.RoleTypes
        End Get
    End Property
    Public Sub RoleTypeSet(ByVal TypeCode As String, _
        ByVal Description As String, _
        Optional ByVal SortOrder As Byte = 0, _
        Optional ByVal Comment As String = "")
        'adds or updates a Role Type record
        _Types.RoleTypeSet(TypeCode, Description, SortOrder, Comment)
    End Sub
    Public Function RoleTypeSelect(ByVal TypeCode As String) As DataRow
        'returns a Role Type record
        Return _TypeSelect(ROLE_TYPE, TypeCode)
    End Function
    Public Sub RoleTypeDelete(ByVal TypeCode As String)
        'deletes a Role Type record
        _Types.RoleTypeDelete(TypeCode)
    End Sub

    '*****************************************************
    ' ServiceTypes 
    '*****************************************************
    Public ReadOnly Property ServiceTypes() As DataTable
        Get
            'lists all Service Types
            Return _Types.ServiceTypes
        End Get
    End Property
    Public Sub ServiceTypeSet(ByVal TypeCode As String, _
        ByVal Description As String, _
        Optional ByVal SortOrder As Byte = 0, _
        Optional ByVal Comment As String = "")
        'adds or updates an Service Type
        _Types.ServiceTypeSet(TypeCode, Description, SortOrder, Comment)
    End Sub
    Public Function ServiceTypeSelect(ByVal TypeCode As String) As DataRow
        'returns an Service Type row
        Return _TypeSelect(SERVICE_TYPE, TypeCode)
    End Function
    Public Sub ServiceTypeDelete(ByVal TypeCode As String)
        'deletes an Service Type row
        _Types.ServiceTypeDelete(TypeCode)
    End Sub
    Public ReadOnly Property OrderStatusTypes() As DataTable
        Get
            'lists all OrderStatus Types; used by Ordermanagement to display available status values for an order
            If _AccountRow("TypeCode") = ASi.AccountBO.Constants.ACCOUNT_TYPECODE_CUSTOMER Then
                'this instance is a Customer so get the parent (merchant) OrderStatusTypes
                Dim t As New AccountTypes(_ConnectString, _ParentID)
                t.Refresh()
                Return t.OrderStatusTypes
            Else
                'return own account's order status values
                Return _Types.OrderStatusTypes
            End If
        End Get
    End Property
    Public ReadOnly Property PackingListStatusTypes() As DataTable
        Get
            Return _Types.PackingListStatusTypes
        End Get
    End Property
    Public ReadOnly Property POStatusTypes() As DataTable
        Get
            Return _Types.POStatusTypes
        End Get
    End Property

#End Region

#Region "Private"
    'private methods and properties
    Private Function _GetAccountDataset(ByVal ParentID As Integer, ByVal AccountID As Integer) As DataSet
        'return an account dataset
        Return ExecuteDataset(_ConnectString, CommandType.StoredProcedure, "AccountSelect", _
                              New SqlClient.SqlParameter("@ParentID", ParentID), _
                              New SqlClient.SqlParameter("@AccountID", AccountID))

    End Function

    Private Sub _RefreshList(ByVal TableOrdinal As AccountTableOrdinals, ByVal RefreshDataSet As Boolean)
        'refresh this object's tables and optionally the complete dataset 
        If RefreshDataSet = True Then
            _AccountDS = _GetAccountDataset(_ParentID, _AccountID)
        End If
        Select Case TableOrdinal
            Case AccountTableOrdinals.Address
                _AddressList = _AccountDS.Tables(AccountTableOrdinals.Address)
            Case AccountTableOrdinals.Phone
                _PhoneList = _AccountDS.Tables(AccountTableOrdinals.Phone)
            Case AccountTableOrdinals.Person
                _PersonList = _AccountDS.Tables(AccountTableOrdinals.Person)
            Case AccountTableOrdinals.Email
                _EmailList = _AccountDS.Tables(AccountTableOrdinals.Email)
                'Case AccountTableOrdinals.ChildAccounts
                '    _ChildList = _AccountDS.Tables(AccountTableOrdinals.ChildAccounts)
            Case AccountTableOrdinals.Users
                _UserList = _AccountDS.Tables(AccountTableOrdinals.Users)
            Case AccountTableOrdinals.Notes
                _NoteList = _AccountDS.Tables(AccountTableOrdinals.Notes)
            Case AccountTableOrdinals.Base
                If _AccountDS.Tables(AccountTableOrdinals.Base).Rows.Count > 0 Then
                    _AccountRow = _AccountDS.Tables(AccountTableOrdinals.Base).Rows(0)
                Else
                    Throw New Exception("There is no account row.")
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
    Private Function _TypeSelect(ByVal Type As String, ByVal TypeCode As String) As DataRow
        'only returns the first instance of a type. This won't suffice for multiple addresses, emails, etc.
        'note to self: see AddressSearch, EmailSearch, etc.
        Type = Type.Trim

        If Type = "" Then
            Throw New Exception("Type identifier is required.")
        End If

        TypeCode = Left(TypeCode.Trim, 20)
        If TypeCode = "" Then
            Throw New Exception(Type & " Type is required.")
        End If

        Dim r As DataRow = Nothing
        Try
            Select Case Type
                Case ACCOUNT_TYPE
                    r = _Types.AccountTypes.Select("Type='" & Type & "' and Code='" & TypeCode & "'")(0)
                Case ADDRESS_TYPE
                    r = _Types.AddressTypes.Select("Type='" & Type & "' and Code='" & TypeCode & "'")(0)
                Case EMAIL_TYPE
                    r = _Types.EmailTypes.Select("Type='" & Type & "' and Code='" & TypeCode & "'")(0)
                Case PERSON_TYPE
                    r = _Types.PersonTypes.Select("Type='" & Type & "' and Code='" & TypeCode & "'")(0)
                Case PHONE_TYPE
                    r = _Types.PhoneTypes.Select("Type='" & Type & "' and Code='" & TypeCode & "'")(0)
                Case ROLE_TYPE
                    r = _Types.RoleTypes.Select("Type='" & Type & "' and Code='" & TypeCode & "'")(0)
                Case SERVICE_TYPE
                    r = _Types.ServiceTypes.Select("Type='" & Type & "' and Code='" & TypeCode & "'")(0)
                Case ORDER_STATUS_TYPE
                    r = _Types.OrderStatusTypes.Select("Type='" & Type & "' and Code='" & TypeCode & "'")(0)
                Case Else
                    Throw New Exception("Unrecognized Type in Private _TypeSelect (" & Type & ")")
            End Select
        Catch ex As Exception
            _LogEvent("_TypeSelect", "AccountID: " & _AccountID.ToString & " Type: " & Type & " TypeCode:" & TypeCode & " Error:" & ex.Message & IIf(Not ex.InnerException Is Nothing AndAlso ex.InnerException.Message <> "", " (Inner: " & ex.InnerException.Message & ")", ""), MessageType.Error)
        Finally
            '_Types.EmailTypes.DefaultView.RowFilter = ""
        End Try

        'return row - could be nothing
        Return r
    End Function
#End Region

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class
Public NotInheritable Class ConfigSettings
    Inherits BaseConfigSettings

    Private Const APP_NAME As String = "ASi.AccountBO"
    Private Const PROCESS As String = "ConfigSettings"

    Private _ConnectString As String = ""
    Private _AccountID As Integer = 0

    Public ReadOnly Property ConnectString() As String
        Get
            Return _ConnectString
        End Get
    End Property
    Public ReadOnly Property AccountID() As Integer
        Get
            Return _AccountID
        End Get
    End Property

    Public Sub New(ByVal ConnectString As String, ByVal AccountID As Integer)

        MyBase.New(ConnectString, AccountID, ACCOUNT_TYPE)

        _ConnectString = ConnectString
        _AccountID = AccountID
        
        'MyBase.RefreshDataset()
        
    End Sub
    Public Sub New(ByVal ConnectString As String, ByVal AccountID As Integer, ByVal Process As String)

        MyBase.New(ConnectString, AccountID, Process)

        _ConnectString = ConnectString
        _AccountID = AccountID

        'MyBase.RefreshDataset()

    End Sub

End Class

