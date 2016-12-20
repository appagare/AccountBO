#Region "Notes"
'   Each instance of this class represents a service (cart, inventory, whatever)
'   A "service" consists of a single entry in the AccountTypeCode table with a Type of Service
'   and a Dataset of configuration values for the service.
'   Each Dataset contains one or more data tables exposed through this class
'   Each table is named as the Group value
'   Each table is exposed as a hashtable through this class
'   Each hashtable is referenced by table name (which is Group value)

'tasks:
'fetch dataset based on code
'loop thru dataset creating hashtables
'how to identify

#End Region
Imports ASi.DataAccess.SqlHelper
Imports ASi.LogEvent.LogEvent
Public Class Service
    Inherits BaseConfigSettings

    Private Const APP_NAME As String = "ASi.AccountBO"
    Private Const PROCESS As String = "Service"
    Private _LogEventBO As ASi.LogEvent.LogEvent

    Private _ConnectString As String = ""
    Private _AccountID As Integer = 0
    Private _ServiceCode As String = ""

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
    Public ReadOnly Property ServiceCode() As String
        Get
            Return _ServiceCode
        End Get
    End Property

    Public Sub New(ByVal ConnectString As String, ByVal AccountID As Integer, ByVal ServiceCode As String)

        MyBase.New(ConnectString, AccountID, ServiceCode)

        _ConnectString = ConnectString
        _AccountID = AccountID
        _ServiceCode = ServiceCode

        Try
            _LogEventBO = New ASi.LogEvent.LogEvent
        Catch ex As Exception
            'consider throwing this error
            'if not, component may work w/out logging
            'if so, component will fail
        End Try

        'MyBase.RefreshDataset()


    End Sub
#Region "Private"
    

    Private Sub _LogEvent(ByVal Src As String, ByVal Msg As String, ByVal Type As ASi.LogEvent.LogEvent.MessageType)
        Try
            _LogEventBO.LogEvent(APP_NAME, Src, Msg, Type, LogType.Queue)
        Catch ex As Exception
            _LogEventBO.LogEvent(APP_NAME, Src, Msg, Type, LogType.SystemEventLog)
        End Try
    End Sub
#End Region

End Class
