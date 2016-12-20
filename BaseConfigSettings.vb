'todo: code this as it is being used.

#Region "Notes"
'Purpose - provide a convenient and fast method of programmatically accessing an accounts configuration settings.
'This can either contain all config settings or it can be inherited to contain a single group
'
'Terminology:
' Process = Top level grouping such as "Account" or service ("Cart", "Order Manager", etc.) - 50
' Group = grouping within Process such as "Display", "Navigation", "Shipping",  etc. - 50
' Parameter = actual parameter name - 100
' Value = parameter value - 7600
' see _HashTables declaration

#End Region

Imports ASi.DataAccess.SqlHelper
Imports ASi.AccountBO.Constants

Public Enum SystemConfigGroups
    All = 0
    MASTER = 1
    SYSTEM = 2
End Enum

Public MustInherit Class BaseConfigSettings
    Private _ConnectString As String = ""
    Private _AccountID As Integer = 0
    Private _Process As String = ""
    Private _DT As New DataTable  'cached dataset
    Private _SystemConfigDV As New DataView 'cached dataview of systemconfig table
    Private _HashTables As Hashtable 'hashtable collection of hashtables - each item is <hashtable_name><hashtable object>

    'NOTE TO SELF - Attribute ADD/DELETE/LIST are duplicated in InventoryManagerBO.Config since it is also used there
    Public Function AttributeList(Optional ByVal VisibleOnly As Boolean = True) As DataView
        Dim dv As New DataView(ASi.DataAccess.SqlHelper.ExecuteDataset(_ConnectString, _
                        CommandType.StoredProcedure, "AttributeList", _
                       New SqlClient.SqlParameter("@AccountID", _AccountID), _
                       New SqlClient.SqlParameter("@Mode", -1)).Tables(0))
        If VisibleOnly = True Then
            dv.RowFilter = "Mode=1"
        Else
            dv.RowFilter = ""
        End If
        Return dv
    End Function
    'Public Function AttributeList() As DataSet
    '    Return ASi.DataAccess.SqlHelper.ExecuteDataset(_ConnectString, _
    '                    CommandType.StoredProcedure, "AttributeList", _
    '                   New SqlClient.SqlParameter("@AccountID", _AccountID), _
    '                   New SqlClient.SqlParameter("@Mode", -1))
    'End Function
    Public Sub AttributeListAdd(ByVal AttributeName As String, Optional ByVal Mode As Integer = 1, Optional DisplayOrder As Integer = 100)
        ExecuteNonQuery(_ConnectString, CommandType.StoredProcedure, "AttributeAdd", _
                New SqlClient.SqlParameter("@AccountID", _AccountID), _
                New SqlClient.SqlParameter("@AttributeName", AttributeName), _
                New SqlClient.SqlParameter("@Mode", Mode), _
                New SqlClient.SqlParameter("@DisplayOrder", DisplayOrder))
    End Sub
    Public Sub AttributeListDelete(ByVal AttributeID As Integer)
        'not sure why this would be necessary here; user can hide an attribute but prevent it from being re-learned by marking it 0
        ExecuteNonQuery(_ConnectString, CommandType.StoredProcedure, "AttributeDelete", _
                New SqlClient.SqlParameter("@AccountID", _AccountID), _
                New SqlClient.SqlParameter("@AttributeID", AttributeID))
    End Sub
    Public Sub AttributeListUpdate(ByVal AttributeID As Integer, ByVal Mode As Byte, Optional ByVal DisplayOrder As Integer = -1)
        ExecuteNonQuery(_ConnectString, CommandType.StoredProcedure, "AttributeUpdate", _
                New SqlClient.SqlParameter("@AccountID", _AccountID), _
                New SqlClient.SqlParameter("@AttributeID", AttributeID), _
                New SqlClient.SqlParameter("@Mode", Mode), _
                New SqlClient.SqlParameter("@DisplayOrder", DisplayOrder))
    End Sub
    'return all rows for an account
    Public Function GetConfigSettings() As DataTable
        Return _DT
    End Function

    'return rows within a Group
    Public Function GetConfigRows(ByVal Group As String) As DataRow()
        'Dim r() As DataRow = _DT.Select("Group = '" & Group & "'", "SortOrder Asc")
        Return _DT.Select("Group = '" & Group & "'", "SortOrder Asc, Parameter Asc")
    End Function
    'return a single row
    Public Function GetConfigRow(ByVal Group As String, ByVal Parameter As String) As DataRow
        Try
            Return _DT.Select("Group = '" & Group & "' and Parameter='" & Parameter & "'")(0)
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    'return rows within a Group as a HashTable
    Public Function GetConfigHashTable(ByVal Group As String) As Hashtable
        Dim h As New Hashtable
        If _HashTables.ContainsKey(Group.ToLower) = True Then
            h = CType(_HashTables(Group.ToLower), Hashtable)
        End If
        Return h
    End Function
    'returns a value from a Group and Parameter from the hashtable
    Public Function GetConfigValue(ByVal Group As String, ByVal Parameter As String) As String
        Dim ReturnValue As String = ""

        If _HashTables.ContainsKey(Group.ToLower) = True Then
            Dim h As New Hashtable
            h = CType(_HashTables(Group.ToLower), Hashtable)
            If h.ContainsKey(Parameter.ToLower) = True Then
                ReturnValue = h(Parameter.ToLower).ToString
            End If
            h = Nothing
        End If
        Return ReturnValue
    End Function

    Public Sub ConfigDelete(ByVal Group As String, _
       ByVal Parameter As String)

        'deletes a record from the Accounts Config table
        ExecuteNonQuery(_ConnectString, _
                    CommandType.StoredProcedure, _
                    "AccountConfigDelete", _
                    New SqlClient.SqlParameter("@AccountID", _AccountID), _
                    New SqlClient.SqlParameter("@Process", _Process), _
                    New SqlClient.SqlParameter("@Group", Group), _
                    New SqlClient.SqlParameter("@Parameter", Parameter))

        'refresh
        RefreshDataset()

    End Sub

    Public Sub ConfigSet(ByVal Group As String, _
        ByVal Parameter As String, _
        ByVal Value As String, _
        Optional ByVal SortOrder As Integer = 0, _
        Optional ByVal Comment As String = "", _
        Optional ByVal CanCopy As Byte = 0)

        'inserts or updates a record into the Accounts Config table
        ExecuteNonQuery(_ConnectString, _
                    CommandType.StoredProcedure, _
                    "AccountConfigSet", _
                    New SqlClient.SqlParameter("@AccountID", _AccountID), _
                    New SqlClient.SqlParameter("@Process", _Process), _
                    New SqlClient.SqlParameter("@Group", Group), _
                    New SqlClient.SqlParameter("@Parameter", Parameter), _
                    New SqlClient.SqlParameter("@Value", Value), _
                    New SqlClient.SqlParameter("@SortOrder", SortOrder), _
                    New SqlClient.SqlParameter("@Comment", Comment), _
                    New SqlClient.SqlParameter("@CanCopy", CanCopy))

        'refresh
        RefreshDataset()

    End Sub

    Public ReadOnly Property ShippingEntriesList() As DataView
        Get
            Return New DataView(_DT, "Group='" & Constants.ACCOUNT_GROUP_PACKING_SLIP & "' and Parameter like 'ShippingEntry%'", "SortOrder asc", DataViewRowState.CurrentRows)
        End Get
    End Property
    Public ReadOnly Property GetPackingSlipNumber As String
        Get
            'this does not lock the table so it could result in gaps and, in rare instances a duplicate
            Dim ReturnValue As String = ""
            Dim fmt As String = ""
            '            'CurrentNumber	1
            'NumberFormat	ABC-[YYYY]-######
            _DT.DefaultView.RowFilter = "Group='" & Constants.ACCOUNT_GROUP_PACKING_SLIP & "' and Parameter='NumberFormat'"
            If _DT.DefaultView.Count > 0 Then
                fmt = _DT.DefaultView(0)("Value")
            End If

            'do the datestamp formatting first
            ReturnValue = Replace(Replace(Replace(Replace(Replace(Replace(fmt, "[YYYY]", Now.Year.ToString, , , CompareMethod.Text), _
                                  "[MM]", Now.Month.ToString, , , CompareMethod.Text), _
                                  "[DD]", Now.Day.ToString, , , CompareMethod.Text), _
                                  "[HH]", Now.Hour.ToString, , , CompareMethod.Text), _
                                  "[NN]", Now.Minute.ToString, , , CompareMethod.Text), _
                                  "[SS]", Now.Second.ToString, , , CompareMethod.Text)

            Dim fmtLength As Integer = 0
            _GetNumberFormatFromFullFormat(fmt, fmtLength)

            'get the number value
            Dim prm As New SqlClient.SqlParameter
            prm.Direction = ParameterDirection.ReturnValue
            Dim i As Integer = ExecuteScalar(_ConnectString, CommandType.StoredProcedure, _
                "AccountGetConfigNumber", _
                New SqlClient.SqlParameter("@AccountID", _AccountID), _
                New SqlClient.SqlParameter("@Process", ACCOUNT_TYPE), _
                New SqlClient.SqlParameter("@Group", ACCOUNT_GROUP_PACKING_SLIP), _
                New SqlClient.SqlParameter("@Parameter", "CurrentNumber"), _
                prm)
            If i = 0 Then
                i = CType(prm.Value, Integer)
            End If

            ReturnValue = Replace(ReturnValue, fmt, i.ToString.PadLeft(fmtLength, "0"))

            Return ReturnValue

        End Get
    End Property
    Private Sub _GetNumberFormatFromFullFormat(ByRef fmt As String, ByRef fmtlen As Integer)
        'find the 'find the "######" part to create a new format string and use VB's "0" which will zero pad it (the # will not pad)
        'format: ABC-[YYYY]-[######]-XYZ
        'spos=12 epos=18
        Dim spos As Integer = InStr(fmt, "[#")
        Dim epos As Integer = InStr(fmt, "#]")
        If epos > spos Then
            fmtlen = (epos - spos)
            fmt = "".PadRight(fmtlen, "#")
        End If
        fmt = "[" & fmt & "]"
    End Sub


    Public Function SystemConfigGet(ByVal Process As SystemConfigGroups, ByVal Group As String, ByVal Parameter As String) As String

        Dim Value As String = ""
        Select Case Process
            Case SystemConfigGroups.MASTER
                _SystemConfigDV.RowFilter = "Process='MASTER' and Group='" & Group & "' and Parameter='" & Parameter & "'"
                If _SystemConfigDV.Count > 0 Then
                    'should be 1 but just get the first entry if more than one
                    Value = _SystemConfigDV(0)("Value")
                End If
            Case SystemConfigGroups.SYSTEM
                _SystemConfigDV.RowFilter = "Process='System' and Group='" & Group & "' and Parameter='" & Parameter & "'"
                If _SystemConfigDV.Count > 0 Then
                    'should be 1 but just get the first entry if more than one
                    Value = _SystemConfigDV(0)("Value")
                End If
            Case Else
                'search all
                _SystemConfigDV.RowFilter = "Group='" & Group & "' and Parameter='" & Parameter & "'"
                If _SystemConfigDV.Count > 0 Then
                    'should be 1 but just get the first entry if more than one
                    Value = _SystemConfigDV(0)("Value")
                End If
        End Select
        'clear filter
        _SystemConfigDV.RowFilter = ""

        Return Value
    End Function


    Public Sub New(ByVal ConnectionString As String, _
                   ByVal AccountID As Integer, _
                   ByVal Process As String)

        _ConnectString = ConnectionString
        _AccountID = AccountID
        _Process = Process

        RefreshDataset()

        'handle SystemConfig settings (instantiate a Dataview since it's readonly so just use RowFilter to get data)
        _SystemConfigDV = ExecuteDataset(_ConnectString, CommandType.StoredProcedure, "SystemConfigList").Tables(0).DefaultView


    End Sub

    Protected Sub RefreshDataset()
        'get the complete dataset
        _DT = ExecuteDataset(_ConnectString, CommandType.StoredProcedure, "AccountConfigSelect", _
                              New SqlClient.SqlParameter("@AccountID", _AccountID), _
                              New SqlClient.SqlParameter("@Process", _Process)).Tables(0)

        'force update of all hashtables 
        'note - considered doing this by Group to minimize unnecessary re-reading if unchanged groups
        'but another user could update a setting; this insures refresh always contains the latest modifications
        RefreshHastTable()

    End Sub
    'protected methods 
    Protected Sub RefreshHastTable()

        Dim ThisName As String = ""
        
        _HashTables = Nothing
        _HashTables = New Hashtable

        'Process, Group, Parameter, Value
        'a Process can have different Groups and each Group has Parameters with Values. Each Group is stored as a separate hashtable within _HashTables
        Dim r As DataRow
        Dim h As New Hashtable

        'for each row, add an entry to the "h" hashtable
        'when group name changes, add the "h" hashtable to the _HashTables hashtable and then start a new "h" hashtable

        For Each r In _DT.Rows
            'create the HashTable and Datatable entry
            If ThisName <> "" And ThisName.ToLower <> CType(r("Group"), String).ToLower Then
                'new group (and not first pass)
                If Not h Is Nothing Then
                    'ThisName is not empty, doesn't match the current Group and "h" is a hashtable;
                    'add "h" to the collection
                    _AddHashTableToHashTables(ThisName, h)
                End If
                h = Nothing
                h = New Hashtable
            End If
            'capture this group name
            ThisName = CType(r("Group"), String).ToLower
            'add the value to the hashtable
            h.Add(CType(r("Parameter"), String).ToLower, CType(r("Value"), String))
        Next

        'handle last group
        If ThisName <> "" AndAlso Not h Is Nothing Then
            'last group
            '
            'add "h" to the collection
            _AddHashTableToHashTables(ThisName, h)
            'done
        End If
        h = Nothing


    End Sub

    Private Sub _AddHashTableToHashTables(ByVal ThisName As String, ByVal h As Hashtable)
        'this chunk of code adds "h" to the collection of hashtables in _HashTables
        If h.Count > 0 Then
            If _HashTables.ContainsKey(ThisName.ToLower) = False Then
                _HashTables.Add(ThisName.ToLower, h)
            Else
                'shouldn't happen but if the _HashTables already contains this, update it
                _HashTables(ThisName.ToLower) = h
            End If
        ElseIf _HashTables.ContainsKey(ThisName.ToLower) = True Then 'h.Count = 0 but HashTable existed
            'don't really anticipate this happening but handle case where where values previously existed but they've now been deleted.
            _HashTables.Remove(ThisName.ToLower)
        End If
    End Sub


End Class
