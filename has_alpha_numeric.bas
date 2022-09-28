Attribute VB_Name = "has_alpha_numeric"
Option Explicit

Public Function funcHasAlphaNumeric(ByVal strTarget As String) As Boolean
    funcHasAlphaNumeric = False
    Dim a$, b$, c$, i As Integer
    'The dollar sign forces the variable to return a string type rather than an undeclared variant.
    'This is faster and this procedure needs to be as fast as it can be.
    a$ = strTarget
    For i = 1 To Len(a$)
        b$ = Mid(a$, i, 1)
        'By default, this does not consider a blank space to be an alphabetic character. If you wan to
        'have this return true on a blank space, then change the following line thus: If b$ Like "[A-Za-z0-9 ]" Then
        If b$ Like "[A-Za-z0-9]" Then
            funcHasAlphaNumeric = True
            Exit Function
        End If
    Next i
End Function

Private Sub TestHasAlphaNumeric()
    'Place your cursor in this procedure and click the play button.
    'Make sure you have the Immediate Window showing (Ctrl + G)
    
    Dim strTest As String
    
    '1. Test for alphabetical: TRUE
    strTest = "State of the Onion Address" 'Hey, I'm an old Perl guy
    Debug.Print "1.  Alphabetic test """ & strTest & """: " & funcHasAlphaNumeric(strTest)
    
    '2. Test for  alphabetical: FALSE
    strTest = ""
    Debug.Print "2a. Alphabetical test """ & strTest & """: " & funcHasAlphaNumeric(strTest)
    
    strTest = " "
    Debug.Print "2b. Alphabetical test """ & strTest & """: " & funcHasAlphaNumeric(strTest)
    
    '3. Test for numeric: TRUE
    strTest = "5067522980"
    Debug.Print "3.  Numeric test """ & strTest & """: " & funcHasAlphaNumeric(strTest)
    
    '4. Test for numerics on an interger or long data types
    Dim intTest As Integer
    intTest = 145
    Debug.Print "4a. Numeric test (integer) """ & intTest & """: " & funcHasAlphaNumeric(intTest)

    Dim longTest As Long
    longTest = 83475082
    Debug.Print "4b. Numeric test (long) """ & longTest & """: " & funcHasAlphaNumeric(longTest)

    '5. Test for alphanumerics in a variant data type
    Dim varTest As Variant
    varTest = "My kung-fu is strong."
    Debug.Print "5a. Alphabetic test (variant) """ & varTest & """: " & funcHasAlphaNumeric(varTest)

    varTest = "Jenny 8675309"
    Debug.Print "5b. Numeric test (variant) """ & varTest & """: " & funcHasAlphaNumeric(varTest)

End Sub
