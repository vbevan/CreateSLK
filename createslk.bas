'---------------------------------------------------------------------------------------
' Module        : CreateSLK
' Author        : Vaughan Bevan (8/04/2015)
' Purpose       : Create Statistical Linkage Key from service users firstname, surname,
'                 date of birth and gender.
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

Function CreateSLK(varSurname As Variant, varFirstname As Variant, varDOB As _
    Variant, varSex As Variant, Optional varCurrentSLK As Variant = "", Optional blnOverwrite As _
    Boolean = False) As String
'---------------------------------------------------------------------------------------
' Procedure     : CreateSLK
' Author        : Vaughan Bevan (9/04/2015)
' Purpose       : Create a service user's Statistical Linkage Key (SLK) using their
'                 Firstname, Surname, Date of Birth and Gender.
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' varSurname    : Service user's surname.
' varFirstname  : Service user's first name.
' varDOB        : Service user's date of birth.
' varSex        : Service user's sex.
' varCurrentSLK : Service's users current SLK. If one of the four above fields is
'                 missing, but CurrentSLK exists, SLK won't be created. Assumption is
'                 those are "SLK Only" records.
' blnOverwrite  : If True, SLKs will be generated for all records. Be careful not to
'                 overwrite existing valid SLKs with Null values in other fields.
'                 Default = False
'
' Usage:
' ~~~~~~
' CreateSLK (Surname, Firstname, DOB, Gender[, CurrentSLK] [,True/False])
'---------------------------------------------------------------------------------------

    Dim strSurname As String
    Dim strFirstname As String
    Dim strDOB As String
    Dim strSex As String
    Dim strCurrentSLK As String
    
    If Not (Nz(varSurname) = "" Or Nz(varFirstname) = "" Or Nz(varDOB) = "" Or Nz(varSex) = "") _
        Or Nz(varCurrentSLK) = "" _
        Or blnOverwrite = True Then
            strSurname = Nz(varSurname)
            strFirstname = Nz(varFirstname)
            strDOB = Nz(varDOB, "01/01/1900")
            Select Case Trim(Nz(varSex, "9"))
                Case "Male", "M", 1
                    strSex = 1
                Case "Female", "F", 2
                    strSex = 2
                Case Else
                    strSex = 9
            End Select
            
            If Len(strDOB) = 9 Then strDOB = "0" & strDOB
            CreateSLK = SLKSurname(strSurname) & SLKFirstname(strFirstname) & strDOB & strSex
        Else
            strCurrentSLK = Nz(varCurrentSLK)
            CreateSLK = strCurrentSLK
    End If

End Function

Function SLKSurname(strSurname As String) As String
    Dim strLastAlpha As String
    strLastAlpha = AlphaNumericOnly(strSurname)

    If Len(strLastAlpha) > 4 Then
        SLKSurname = Mid$(strLastAlpha, 2, 2) & Mid$(strLastAlpha, 5, 1)
    ElseIf Len(strLastAlpha) > 2 Then
        SLKSurname = Mid$(strLastAlpha, 2, 2) & "2"
    ElseIf Len(strLastAlpha) > 1 Then
        SLKSurname = Mid$(strLastAlpha, 2, 1) & "22"
    ElseIf Len(strLastAlpha) = 1 Then
        SLKSurname = "222"
    Else
        SLKSurname = "999"
    End If
    SLKSurname = UCase$(SLKSurname)
End Function

Function SLKFirstname(strFirstname As String) As String
    Dim strFirstAlpha As String
    strFirstAlpha = AlphaNumericOnly(strFirstname)
    
    If Len(strFirstAlpha) > 2 Then
        SLKFirstname = Mid$(strFirstAlpha, 2, 2)
    ElseIf Len(strFirstAlpha) > 1 Then
        SLKFirstname = Mid$(strFirstAlpha, 2, 1) & "2"
    ElseIf Len(strFirstAlpha) = 1 Then
        SLKFirstname = "22"
    Else
        SLKFirstname = "99"
    End If
    SLKFirstname = UCase$(SLKFirstname)
End Function

Function AlphaNumericOnly(strSource As String) As String
    'Purpose:   Remove non-alphanumeric characters from string.
    
    Dim i As Integer
    Dim strResult As String
    strResult = Trim$(strResult)

    For i = 1 To Len(strSource)
        Select Case AscW(Mid$(strSource, i, 1))
        Case 65 To 90, 97 To 122  'include 32 if you want to include space
            strResult = strResult & Mid$(strSource, i, 1)
        End Select
    Next
    AlphaNumericOnly = strResult

End Function