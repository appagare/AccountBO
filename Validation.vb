Imports ASi.AccountBO.Constants
Friend NotInheritable Class Validation
    Private _Utility As New ASi.UtilityHelper.Utilities

    Friend Function ValidAccountParameters(ByRef AccountName As String, _
            ByRef URL As String, _
            ByRef TypeCode As String, _
            ByRef InActiveDate As String) As String

        Dim ReturnValue As String = ""
        'validation
        AccountName = Left(AccountName.Trim, 200)
        URL = Left(URL.Trim, 2000)
        TypeCode = Left(TypeCode.Trim, 20)
        InActiveDate = InActiveDate.Trim
        If TypeCode = "" Then
            TypeCode = Constants.BLANK
        End If
        If AccountName = "" Then
            ReturnValue = "Account Name cannot be empty"
        End If
        'return error string, if any
        Return ReturnValue

    End Function

    Friend Function ValidAddressParameters(ByRef Address1 As String, _
        ByRef Address2 As String, _
        ByRef City As String, _
        ByRef StateCode As String, _
        ByRef PostalCode As String, _
        ByRef TypeCode As String, _
        ByRef CountryCode As String) As String

        'validation
        Address1 = Left(Address1.Trim, 100)
        Address2 = Left(Address2.Trim, 100)
        City = Left(City.Trim, 50)
        StateCode = Left(StateCode.Trim, 50)
        PostalCode = Left(PostalCode.Trim, 20)
        CountryCode = Left(CountryCode.Trim, 50)
        TypeCode = Left(TypeCode.Trim, 20)
        If TypeCode = "" Then
            TypeCode = Constants.BLANK
        End If

        Dim ReturnValue As String = ""

        If Address1 = "" Then
            ReturnValue &= "Address1,"
        End If
        If City = "" Then
            ReturnValue &= "City,"
        End If
        If PostalCode = "" OrElse _Utility.ValidZip(PostalCode) = False Then
            ReturnValue &= "PostalCode,"
        End If

        If ReturnValue <> "" Then
            If Right(ReturnValue, 1) = "," Then
                ReturnValue = Left(ReturnValue, Len(ReturnValue) - 1)
            End If
            ReturnValue = "The following values are either missing or invalid: " & ReturnValue
        End If

        'return error string, if any
        Return ReturnValue

    End Function

    Friend Function ValidEmailParameters(ByRef EmailName As String, _
            ByRef EmailAddress As String, _
            ByRef TypeCode As String) As String

        'validation
        EmailName = Left(EmailName.Trim, 50)
        EmailAddress = Left(EmailAddress.Trim, 200)
        TypeCode = Left(TypeCode.Trim, 20)
        If TypeCode = "" Then
            TypeCode = BLANK
        End If

        Dim ReturnValue As String = ""

        'If EmailName = "" Then
        '    ReturnValue &= "Description,"
        'End If
        If EmailAddress = "" OrElse _Utility.ValidEmail(EmailAddress) = False Then
            ReturnValue &= "EmailAddress,"
        End If

        If ReturnValue <> "" Then
            If Right(ReturnValue, 1) = "," Then
                ReturnValue = Left(ReturnValue, Len(ReturnValue) - 1)
            End If
            ReturnValue = "The following values are either missing or invalid: " & ReturnValue
        End If

        'return error string, if any
        Return ReturnValue

    End Function

    Friend Function ValidPhoneParameters(ByRef PhoneNumber As String, _
            ByRef Extension As String, _
            ByRef TypeCode As String) As String

        'validation
        PhoneNumber = Left(PhoneNumber.Trim, 20)
        Extension = Left(Extension.Trim, 8)
        TypeCode = Left(TypeCode.Trim, 20)
        If TypeCode = "" Then
            TypeCode = BLANK
        End If

        Dim ReturnValue As String = ""

        If PhoneNumber = "" OrElse _Utility.ValidPhone(PhoneNumber) = False Then
            If PhoneNumber = "" Then
                ReturnValue = "Phone Number is required."
            Else
                ReturnValue = "Phone Number format is invalid. Must be NNN-NNN-NNNN or (NNN) NNN-NNNN or NNN.NNN.NNNN."
            End If
        End If

        'return error string, if any
        Return ReturnValue

    End Function

    Friend Function ValidPersonParameters(ByVal FirstName As String, _
       ByVal LastName As String, _
       ByVal TypeCode As String, _
       ByVal Prefix As String, _
       ByVal MI As String, _
       ByVal Suffix As String, _
       Optional ByVal LastNameRequired As Boolean = True) As String

        'validation
        Prefix = Left(Prefix.Trim, 10)
        FirstName = Left(FirstName.Trim, 50)
        MI = Left(MI.Trim, 10)
        LastName = Left(LastName.Trim, 50)
        Suffix = Left(Suffix.Trim, 10)
        TypeCode = Left(TypeCode.Trim, 20)
        If TypeCode = "" Then
            TypeCode = BLANK
        End If

        Dim ReturnValue As String = ""
        If FirstName = "" AndAlso LastName = "" Then
            ReturnValue &= "Full name,"
        End If

        'revisit this in the future; for now require first or last so Cher and Bono can use the system.

        'If FirstName = "" Then
        'ReturnValue &= "First name,"
        'End If
        'If LastName = "" AndAlso LastNameRequired = True Then
        'ReturnValue &= "LastName,"
        'End If

        If ReturnValue <> "" Then
            If Right(ReturnValue, 1) = "," Then
                ReturnValue = Left(ReturnValue, Len(ReturnValue) - 1)
            End If
            ReturnValue = "The following values are either missing or invalid: " & ReturnValue
        End If

        'return error string, if any
        Return ReturnValue

    End Function

    Friend Function ValidNoteParameters(ByRef NoteTitle As String, _
            ByRef Note As String) As String

        Dim ReturnValue As String = ""
        'validation
        NoteTitle = Left(NoteTitle.Trim, 50)
        Note = Left(Note.Trim, 7000)
        If Note = "" Then
            ReturnValue = "Note cannot be empty"
        End If
        'return error string, if any
        Return ReturnValue

    End Function

    Friend Function ValidUserUpdateParameters(ByVal UserName As String, _
                          ByVal EncryptedPassword As String) As String

        'validation
        UserName = UserName.Trim
        EncryptedPassword = EncryptedPassword.Trim

        Dim ReturnValue As String = ""
        If Len(UserName) < MIN_USERNAME_LENGTH Then
            ReturnValue &= "Username must be " & MIN_USERNAME_LENGTH.ToString & " characters or greater."
        ElseIf Len(UserName) > 50 Then
            ReturnValue &= "Username cannot exceed 50 characters in length."
        End If

        'shouldn't happen - plain-text password should already have been validated by this point
        If Len(EncryptedPassword) < 1 Then
            ReturnValue &= "Invalid password."
        End If

        If ReturnValue <> "" Then
            If Right(ReturnValue, 1) = "," Then
                ReturnValue = Left(ReturnValue, Len(ReturnValue) - 1)
            End If
            ReturnValue = "The following values are either missing or invalid: " & ReturnValue
        End If

        'return error string, if any
        Return ReturnValue

    End Function

    Sub New()

    End Sub
End Class
