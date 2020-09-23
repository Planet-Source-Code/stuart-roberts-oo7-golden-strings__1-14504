Attribute VB_Name = "Strings"
Option Explicit

' --------------------------------------------------------------'
' ROUTINE:      CountChars
' DESCRIPTION:
'               Returns the number of times a given character appears
'               in a string.
'               If bGroup is True then sub counts a group of characters
'               as 1. " " is search char; "This is a   Test" is string
'               If bGroup is False returns 5 otherwise 3.
' PARAMETERS:
'       bGroup - If sub is to count groups of characters or not
' --------------------------------------------------------------'
Public Function CountChars(sChar As String, ByVal sString As String, Optional bGroup As Boolean = False) As Integer
'# Returns the number of times a given character occurs in a string
On Error GoTo HandleError
    Dim iCount As Integer, iPos As Integer, iLength As Integer
    
    iPos = 1
    iLength = 1
    CountChars = 0
    If (Len(sString) <= 0) Or (Len(sChar) <= 0) Then Exit Function
    
    Do While iLength <= Len(sString)
        iPos = InStr(Mid$(sString, iLength), sChar)
        If iPos > 0 Then
'# Only increase the count of a character if -
'# the user wants to count grouped characters seperately and
'# the previous character is the same type
            If (bGroup) And (iLength + iPos - 2) >= 1 Then
                If Not (Mid$(sString, iLength + iPos - 2, 1) = Left$(sChar, 1)) Then
                    iCount = iCount + 1
                End If
            Else
'# Count character as above is not true
                iCount = iCount + 1
            End If
'# Always adjust iLength to help move through string
            iLength = iLength + iPos
        Else
            Exit Do
        End If
    Loop
    CountChars = iCount
    Exit Function
    
HandleError:
    CountChars = -1
    HandleErr "Strings.CountChars"
End Function

Public Function DelLeftAfter(sChars As String, ByVal sLine As String, Optional bGroup As Boolean = True) As String
' Removes unwanted characters from left of given string
' # EXAMPLE
'       For each character in sChars function will remove all occurances of it
'       from string. An occurance can be anything from 1 character to
'       a batch of characters.
'       "HERE 45  56"
'       sChars would be "  " (2 spaces) to retrieve "56"
' # RETURNS formatted string
On Error GoTo HandleError
    Dim iCount As Integer, iInLoop As Integer
    Dim sChar As String
    
    DelLeftAfter = ""
' Remove unwanted characters to left of folder name
    For iCount = 1 To Len(sChars)
' Retrieve character from start string to look for in folder string (sLine)
        sChar = Mid$(sChars, iCount, 1)
' Remove all characters to left of found string
        sLine = Mid$(sLine, InStr(sLine, sChar) + 1)
' Then remove all characters of same type directly next to it
        If bGroup Then
            iInLoop = 0
            While Mid$(sLine, iInLoop + 1, 1) = sChar
                iInLoop = iInLoop + 1
            Wend
            If iInLoop > 0 Then
                sLine = Mid$(sLine, iInLoop + 1)
            End If
        End If
    Next iCount
    DelLeftAfter = sLine
    Exit Function
    
HandleError:
    HandleErr "Strings.DelLeftAfter"
End Function

Public Function DelRightAfter(sChars As String, ByVal sLine As String, Optional bGroup As Boolean = True) As String
' Removes unwanted characters from right of given string
' # EXAMPLE
'       For each character in sChars - function will remove all occurances of it
'       from string. An occurance can be anything from 1 character to
'       a batch of characters.
'       "123   23 HERE"
'       sChars would be "  " (2 spaces) to retrieve "123"
' # RETURNS formatted string
On Error GoTo HandleError
    Dim iCount As Integer, iInLoop As Integer
    Dim sChar As String
    
    DelRightAfter = ""
    sLine = ReverseString(sLine)
    sChars = ReverseString(sChars)
    sLine = DelLeftAfter(sChars, sLine, bGroup)
    DelRightAfter = ReverseString(sLine)
    Exit Function
    
HandleError:
    HandleErr "Strings.DelRightAfter"
End Function

Public Function FormatLine(sFromLeft As String, sFromRight As String, ByVal sLine As String, Optional bGroup As Boolean = True) As String
On Error GoTo HandleError
' Calls Left and Right Format functions to remove unwanted characters
' from left and right hand-side of given string.
' # RETURNS formatted string
    Dim sNewLine As String
    
    FormatLine = ""
    sNewLine = Trim(sLine)
    sNewLine = DelLeftAfter(sFromLeft, sNewLine, bGroup)
    sNewLine = DelRightAfter(sFromRight, sNewLine, bGroup)
    
    If Len(sNewLine) > 0 Then
        FormatLine = sNewLine
    End If
    Exit Function
    
HandleError:
    HandleErr "Strings.FormatLine"
End Function

Public Function ReverseString(ByVal sString As String, Optional iSplitBy As Integer = 500) As String
On Error GoTo HandleError
' # RETURNS the reverse of the given string
    Dim iCount As Integer, iLoop As Integer
    Dim sReverse As String, saStrings() As String
    
    ReverseString = ""
    If Len(sString) = 0 Then
        Exit Function
    End If
    
'# Re-dimention saStrings to 1/iSplitBy'th of the length of given string
    ReDim saStrings(Int(Len(sString) \ iSplitBy) + 1)
'# copy each iSplitBy'th segment of given string to each increment of array
    For iLoop = 1 To UBound(saStrings)
        saStrings(iLoop) = Mid$(sString, ((iLoop - 1) * iSplitBy) + 1, iSplitBy)
    Next iLoop
    
'# Reverse each element in array and add to reversed string
    For iLoop = UBound(saStrings) To 1 Step -1
        For iCount = Len(saStrings(iLoop)) To 1 Step -1
            sReverse = sReverse + Mid$(saStrings(iLoop), iCount, 1)
        Next iCount
    Next iLoop
    
    ReverseString = sReverse
    Exit Function
    
HandleError:
    HandleErr "Strings.ReverseString"
End Function

Public Function GetLeft(ByVal sData As String, iLength As Integer) As String
On Error GoTo HandleError
' Returns the first iLength characters from sData
' Dismisses hidden characters
' EXAMPLE sData is "A Bruce" and iLength is 3
'         Returns - "ABr"
    Dim sChar As String, sReturn As String, tmpString As String
    Dim iLoop As Integer
    
    GetLeft = ""
    iLoop = 1
    tmpString = sData
    Do While (Len(sReturn) < iLength) And (tmpString > "") And (iLoop <= Len(tmpString))
        sChar = Mid$(tmpString, iLoop, 1)
        If (Asc(sChar) >= 33) And (Asc(sChar) <= 126) Then
            sReturn = sReturn & sChar
        End If
        iLoop = iLoop + 1
    Loop
    
    GetLeft = sReturn
    Exit Function
    
HandleError:
    HandleErr "Strings.GetLeft"
End Function

Public Function GetRight(ByVal sData As String, iLength As Integer) As String
On Error GoTo HandleError
' Returns the last iLength characters from sData
' Dismisses hidden characters
' EXAMPLE sData is "John S" and iLength is 3
'         Returns - "hnS"
    Dim tmpString As String
    
    GetRight = ""
    tmpString = ReverseString(sData)
    tmpString = GetLeft(tmpString, iLength)
    tmpString = ReverseString(tmpString)
    
    GetRight = tmpString
    Exit Function
    
HandleError:
    HandleErr "Strings.GetRight"
End Function

Public Function DelInvisible(ByVal sLine As String, _
                             Optional bDelSpaces As Boolean = False) As String
On Error GoTo HandleError
' Remove all in-visible characters (except Space).
' Accept only ASCii values between 32 and 126
    Dim iCount As Integer, nStart As Integer
    Dim sFormatted As String
    
    DelInvisible = ""
    nStart = 32
    If bDelSpaces Then nStart = 33
    
    For iCount = 1 To Len(sLine)
        Select Case Asc(Mid$(sLine, iCount, 1))
            Case nStart To 126
                sFormatted = sFormatted + Mid$(sLine, iCount, 1)
        End Select
    Next iCount
    DelInvisible = sFormatted
    Exit Function
    
HandleError:
    HandleErr "Strings.DelInvisible"
End Function

Public Function ConvertToUpperCase(ByVal strIn As String, Optional iPos As Integer = 1) As String
'# Converts iPos character of string parameter to uppercase
On Error GoTo HandleError
    Dim strUCase As String
    Dim strPart As String
    
    ConvertToUpperCase = ""
    If Len(strIn) < iPos Then
        Exit Function
    End If
    strPart = strIn
    strUCase = UCase(Mid$(strIn, iPos, 1))
    Mid$(strPart, iPos, Len(strPart)) = strUCase
    ConvertToUpperCase = strPart
    
    Exit Function
    
HandleError:
    If Err.Number = 5 Then
        Resume Next
    End If
    HandleErr "Strings.ConvertToUpperCase"
End Function

Public Function BuildTabbedString(ByVal lstStrings As Control, Optional sBreakChar As String = vbTab) As String
'# Creates a tab deliminated string from a given list control
On Error GoTo HandleError
    Dim sString As String
    Dim iLoop As Integer
    
    BuildTabbedString = ""
    
    If lstStrings.ListCount <= 0 Then Exit Function
    sString = lstStrings.List(0)

    For iLoop = 1 To lstStrings.ListCount - 1
        sString = sString & sBreakChar & lstStrings.List(iLoop)
    Next iLoop
    
    BuildTabbedString = sString
    Exit Function
    
HandleError:
    HandleErr "Strings.BuildTabbedString"
End Function

Public Function ListItemToString(ByVal lstItem As Object, Optional sBreakChar As String = vbTab, _
                                 Optional bChecked As Boolean = False) As String
On Error GoTo HandleError
    Dim iLoop As Integer
    Dim sString As String
    
    ListItemToString = ""
    If bChecked Then
        If lstItem.Checked Then
            sString = "TRUE"
        Else
            sString = "FALSE"
        End If
    Else
        sString = lstItem.Text
    End If
    
    For iLoop = 1 To lstItem.ListSubItems.Count
        sString = sString & sBreakChar & lstItem.SubItems(iLoop)
    Next
    ListItemToString = sString
    
    Exit Function
    
HandleError:
    HandleErr "Strings.ListItemToString"
End Function

Public Function StringToList(ByVal sText As String, lstControl As Control, Optional sChar As String = vbTab) As Boolean
'# Convert a string seperated by special characters to a list box
On Error GoTo HandleError
    Dim sString As String
    Dim iLoop As Integer, iCount As Integer
    
    StringToList = False
    If Len(sText) <= 0 Then Exit Function
    
    iCount = CountChars(sChar, sText) + 1
    
    For iLoop = 1 To iCount
        lstControl.AddItem RetrieveString(sText, sChar, iLoop, False, True)
    Next iLoop
    
    StringToList = True
    Exit Function
    
HandleError:
    HandleErr "Strings.StringToList"
End Function

Public Function StringStarts(ByVal sString As String, sStart As String) As Integer
'# Determines if a string starts with the same characters as sStart string
'# RETURNS - Number of characters in matched string
'          - Or zero if no match
On Error GoTo HandleError
    Dim sCompare As String
    
    StringStarts = 0
    If Len(sString) >= Len(sStart) Then
        sCompare = Left$(sString, Len(sStart))
        If UCase(sCompare) = UCase(sStart) Then
            StringStarts = Len(sCompare)
        End If
    End If
    Exit Function
    
HandleError:
    HandleErr "String.StringStarts"
End Function

Public Function StringEnds(ByVal sString As String, sEnds As String) As Integer
'# Determines if a string ends with the same characters as sEnds string
'# RETURNS - Number of characters in matched string
'          - Or zero if no match
On Error GoTo HandleError
    
    StringEnds = 0
    If Len(sString) >= Len(sEnds) Then
'# Reverse both strings then call StringStarts function
        sEnds = ReverseString(sEnds)
        sString = ReverseString(sString)
        StringEnds = StringStarts(sString, sEnds)
    End If
    Exit Function
    
HandleError:
    HandleErr "String.StringEnds"
End Function

Public Function RetrieveString(ByVal sString As String, Optional sSeperate As String = vbTab, _
                               Optional iNumber As Integer = 1, Optional bGroup As Boolean = True, _
                               Optional bReturnStringOnError As Boolean = False) As String
'# Retrieves the iNumber'th part of a tabbed string
'# If the strings was "One<tab>Two<tab>Three" and inumber was 2
'# would return "Two"
'# If the requested part doesn't exist will return a blank string
On Error GoTo HandleError
    Dim iCharCount As Integer, iPos As Integer
    
    If bReturnStringOnError Then
        RetrieveString = sString
    Else
        RetrieveString = ""
    End If
    If Len(sSeperate) = 0 Then Exit Function
'# Number of seperate strings is iCharCount + 1
    iCharCount = CountChars(Left$(sSeperate, 1), sString, bGroup)
    If (iCharCount <= 0) Or (iNumber > (iCharCount + 1)) Or (iNumber <= 0) Then Exit Function
    
'# Delete the text to the right of and including the seperator
'# character.  DelRightAfter reads the given string backwards!
    RetrieveString = DelRightAfter(String(iCharCount - iNumber + 1, sSeperate), sString, bGroup)
'# Delete all characters to the left of and including the
'# seperator character.
    RetrieveString = DelLeftAfter(String(iCharCount, sSeperate), RetrieveString, bGroup)
    Exit Function
    
HandleError:
    HandleErr "String.RetrieveString"
End Function

Public Sub IfStrBlankAssign(sString As String, ByRef sValue As String)
On Error GoTo HandleError
    
    If Len(sString) <= 0 Then
        sString = sValue
    End If
    Exit Sub
    
HandleError:
    HandleErr "Strings.IfStrBlankAssign"
End Sub

Public Function ChangeChars(ByVal sString As String, sFrom As String, _
                            sTo As String) As String
On Error GoTo HandleError
    Dim nLoop As Integer, nPos As Integer, nLast As Integer
    
    ChangeChars = sString
    If sString = "" Then Exit Function
    If sFrom = "" Then Exit Function
    If sTo = "" Then Exit Function
    
    nPos = InStr(1, sString, sFrom, vbTextCompare)
    If nPos <= 0 Then Exit Function
    nLast = nPos
    While nPos > 0
        sString = Left$(sString, nPos - 1) & sTo & Mid$(sString, nPos + Len(sFrom))
        nPos = InStr(1, sString, sFrom, vbTextCompare)
        If nPos <= nLast Then nPos = -1
    Wend
    
    ChangeChars = sString
    
    Exit Function
    
HandleError:
    HandleErr "Strings.ChangeChars"
End Function
