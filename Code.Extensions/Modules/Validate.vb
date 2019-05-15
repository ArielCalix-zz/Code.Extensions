Imports System.Runtime.CompilerServices
Imports System.Windows.Forms

Public Module Validate
    <Extension()>
    Function Digits(ByVal Input As KeyPressEventArgs) As KeyPressEventArgs
        If IsSpecialKeys(Input) = False Then
            If Char.IsDigit(Input.KeyChar) Then
                Input.Handled = False
            Else
                Input.Handled = True
            End If
        End If
        Return Input
    End Function
    <Extension()>
    Function Letters(ByVal Input As KeyPressEventArgs) As KeyPressEventArgs
        If IsSpecialKeys(Input) = False Then
            If Char.IsLetterOrDigit(Input.KeyChar) Then
                Input.Handled = False
            Else
                Input.Handled = True
            End If
        End If
        Return Input
    End Function
    <Extension()>
    Function LettersAndDigits(ByVal Input As KeyPressEventArgs) As KeyPressEventArgs
        If IsSpecialKeys(Input) = False Then
            If Char.IsLetterOrDigit(Input.KeyChar) Then
                Input.Handled = False
            Else
                Input.Handled = True
            End If
        End If
        Return Input
    End Function
    <Extension()>
    Function Email(ByVal Input As KeyPressEventArgs) As KeyPressEventArgs
        If IsSpecialKeys(Input) = False Then
            If Char.IsLetterOrDigit(Input.KeyChar) Then
                Input.Handled = False
            ElseIf Asc(Input.KeyChar) = 64 Then
                Input.Handled = True
            End If
        End If
        Return Input
    End Function
    <Extension()>
    Function ProductId(ByVal Input As KeyPressEventArgs, ByRef EnableNumbers As Boolean) As KeyPressEventArgs
        If IsSpecialKeys(Input) = False Then
            If EnableNumbers = True Then
                If Char.IsLetterOrDigit(Input.KeyChar) Then
                    Input.Handled = False
                Else
                    Input.Handled = True
                End If
            Else
                If Char.IsLetter(Input.KeyChar) Then
                    Input.Handled = False
                Else
                    Input.Handled = True
                End If
            End If
        End If
        Return Input
    End Function
    <Extension()>
    Function Name(ByVal Input As KeyPressEventArgs, ByRef EnableSpaces As Boolean)
        If IsSpecialKeys(Input) = False Then
            If EnableSpaces = True Then
                If Char.IsLetter(Input.KeyChar) Or Char.IsSeparator(Input.KeyChar) Then
                    Input.Handled = False
                Else
                    Input.Handled = True
                End If
            Else
                If Char.IsLetter(Input.KeyChar) Then
                    Input.Handled = False
                Else
                    Input.Handled = True
                End If
            End If
        End If
        Return Input
    End Function
    <Extension()>
    Function Description(ByVal Input As KeyPressEventArgs)
        If IsSpecialKeys(Input) = False Then
            If Char.IsLetterOrDigit(Input.KeyChar) Or Char.IsSeparator(Input.KeyChar) Or Input.KeyChar = "." Or
                Input.KeyChar = "," Or Input.KeyChar = "/" Or Input.KeyChar = "-" Then
                Input.Handled = False
            Else
                Input.Handled = True
            End If
        End If
        Return Input
    End Function
    <Extension()>
    Function Price(ByVal Input As KeyPressEventArgs, ByRef HasDot As Boolean)
        If IsSpecialKeys(Input) = False Then
            If HasDot = False Then
                If Input.KeyChar = "." Or Char.IsNumber(Input.KeyChar) Then
                    Input.Handled = False
                Else
                    Input.Handled = True
                End If
            Else
                If Char.IsNumber(Input.KeyChar) Then
                    Input.Handled = False
                Else
                    Input.Handled = True
                End If
            End If
        End If
        Return Input
    End Function
    Function IsSpecialKeys(ByVal Input As KeyPressEventArgs) As Boolean
        If Asc(Input.KeyChar) = 8 Or Asc(Input.KeyChar) = 13 Or Input.KeyChar = vbBack Then
            Input.Handled = False
            Return True
        Else
            Return False
        End If
    End Function
End Module