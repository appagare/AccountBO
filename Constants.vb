'class containing useful constants shared by other classes
Public Class Constants
    'account record types
    Public Const ACCOUNT_TYPE As String = "Account"
    Public Const ADDRESS_TYPE As String = "Address"
    Public Const PHONE_TYPE As String = "Phone"
    Public Const PERSON_TYPE As String = "Person"
    Public Const EMAIL_TYPE As String = "Email"
    Public Const ROLE_TYPE As String = "Role"
    Public Const SERVICE_TYPE As String = "Service"

    Public Const PACKING_SLIP_STATUS_TYPE As String = "PACKSLIP_STATUS" 'this is in AccountTypeCode and each entry should correspond with an entry in Datastore StatusLookup
    Public Const PO_STATUS_TYPE As String = "PO_STATUS" 'this is in AccountTypeCode and each entry should correspond with an entry in Datastore StatusLookup

    Public Const ORDER_STATUS_TYPE As String = "ORDER_STATUS" 'this is in AccountTypeCode and each entry should correspond with an entry in Datastore StatusLookup
    Public Const ORDER_STATUS_CODE As String = "ORDER" 'this is in Datastore StatusLookup
    Public Const ITEM_STATUS_CODE As String = "ITEM" 'this is in Datastore StatusLookup

    Public Const BLANK As String = "_BLANK"
    Public Const ACCOUNT_TYPECODE_CUSTOMER As String = "CUSTOMER"
    Public Const ACCOUNT_TYPECODE_MERCHANT As String = "MERCHANT"
    Public Const ACCOUNT_TYPECODE_SUPPLIER As String = "SUPPLIER"

    Public Const ADDRESS_TYPECODE_BILLING As String = "BILLING"
    Public Const ADDRESS_TYPECODE_SHIPPING As String = "SHIPPING"

    'add values here as necessary for config settings/preferences
    Public Const ACCOUNT_GROUP_EMAIL As String = "Email"
    Public Const ACCOUNT_GROUP_SYSTEM As String = "System"
    Public Const ACCOUNT_GROUP_PACKING_SLIP As String = "PACKSLIP"
    Public Const ACCOUNT_GROUP_PURCHASE_ORDER As String = "PO"

    'these are also defined in ASecureCart.OrderManamentBO.Constants so if you change them here, change them there, too
    Public Const ACCOUNT_GROUP_USPS As String = "USPS"
    Public Const ACCOUNT_GROUP_UPS As String = "UPS"
    Public Const ACCOUNT_GROUP_FEDEX As String = "FedEx"

    Public Const INVENTORY_PARAMETER As String = "Inventory"
    Friend Const MIN_USERNAME_LENGTH As Integer = 4

    'table ordinals used by the AccountSelect SP
    Friend Enum AccountTableOrdinals
        Base = 0
        Address = 1
        Phone = 2
        Email = 3
        Person = 4
        Users = 5
        Notes = 6
        'Service = 7
        'OrderStatus = 7
        'ChildAccounts = 5 'removed 12-2-2012
    End Enum
    'table ordinals used by the PersonSelect SP
    Friend Enum PersonTableOrdinals
        Base = 0
        Address = 1
        Phone = 2
        Email = 3
        Users = 4
        UserRoles = 5
    End Enum

    'table ordinals used by the UserSelect SP
    Friend Enum UserTableOrdinals
        Base = 0
        UserRoles = 1
    End Enum

    'config groups
    'Friend Const ACCOUNT_GROUP As String = "Account"
    'Friend Const SYSTEM_GROUP As String = "System"
    'Public Enum ConfigurationGroups
    'Account = 0
    'System = 1
    'End Enum

    'not sure if I like this hard-coded or not...
    Public Enum AccountStatus
        Pending = 0
        Active = 1
        InActive = 2
        Invalid = 3
        NewLock = 4
        AdminLocked = 5
        Unused6 = 6
        Unused7 = 7
        Unused8 = 8
        Deleted = 9
    End Enum
    'user login only - new accounts use NewLock at the account level
    Public Enum UserStatus
        AdminLock = -1
        Locked = 0
        Active = 1
    End Enum
    Public Enum TokenState
        Locked = -1
        Invalid = 0
        Valid = 1
        Expired = 2
    End Enum
    Private Sub New()

    End Sub
End Class
