Attribute VB_Name = "basGeneral"
' Copyright 2000, 2001 Justin Casey (Mersault).
' This file is part of Actor.
'
'    Actor is free software; you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation; either version 2 of the License, or
'    (at your option) any later version.
'
'    Actor is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with Actor; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA

' *** General functions by me for use in any programs
' *** July 99

Option Explicit

' FindLast: Find the last occurence of a string within a string
' Trust: Doesn't check strings are not null
' Arguments: strToSearch, strToFind
' Returns: long - 0 if not found
Public Function FindLast(strToSearch As String, strToFind As String) As Long
    
    Dim lPos As Long
    
    ' keep looking through the string until we go past the last strToFind in strToSearch
    Do
        FindLast = lPos
        lPos = InStr(lPos + 1, strToSearch, strToFind)
    Loop Until lPos = 0
    
End Function

' GetFileTitle: Return the file title for any filename given
' Trust:
' Arguments: strFileName
' Returns: string
Public Function GetFileTitle(strFileName As String) As String

    GetFileTitle = Right$(strFileName, Len(strFileName) - FindLast(strFileName, "\"))
    
End Function

' GetPath: Return the path of any filename given
' Trust:
' Arguments: strFileName
' Returns: string
Public Function GetPath(strFileName As String) As String

    GetPath = Left$(strFileName, FindLast(strFileName, "\"))

End Function

' GetVersion: Construct an ordinal version number out of the three point version number
' Trust: Doesn't check point version numbers
' Arguments: nMajor, nMinor, nRevision
' Returns: integer - ordinal version number
Public Function GetVersion(ByVal nMajor As Integer, ByVal nMinor As Integer, ByVal nRevision As Integer) As Integer

    GetVersion = nMajor * 100 + nMinor * 10 + nRevision

End Function

' InputLine: Replacement for the Input Line # function which can deal with text files
' InputLine: where the lines are delimited by &H0A (a la UNIX) rather than &H0D or &H0A0D
' Trust: Assume nFileNo is a valid open file
' Arguments: nFileNo
' Returns: String read from input minus delimiting characters
' Bugs: Side effect that the majority of files read (CRLF) will return an empty sentence (because of the two characters) for every sentence read
Public Function InputLine(nFileNo As Integer) As String

    Dim strInputChar As String, strNull As String
    
    InputLine = ""
    
    ' read from the file until some delimiting character or the end of the file is encountered
    Do
        If EOF(nFileNo) Then Exit Function
        
        strInputChar = Input(1, nFileNo)
        
        ' If we encounter a carriage return then assume we're dealing with a CRLF delimited file and discard the LF
        If strInputChar = vbCr Then strNull = Input(1, nFileNo): Exit Do
        ' Otherwise just throw away the LF (UNIX)
        If strInputChar = vbLf Then Exit Do
        
        ' add the last character onto the string that will be returned
        InputLine = InputLine & strInputChar
    Loop

End Function

' RemoveQuotes: Remove quotes from the start and end of the string given, if they exist
' Trust:
' Arguments: strBefore
' Returns: string - string with any quotes removed
Public Function RemoveQuotes(strBefore As String) As String

    RemoveQuotes = strBefore
    
    ' Remove quotes if they appear at the start or the end of the given string
    If Asc(Left(strBefore, 1)) = &H22 Then RemoveQuotes = Right$(RemoveQuotes, Len(RemoveQuotes) - 1)
    If Asc(Right$(RemoveQuotes, 1)) = &H22 Then RemoveQuotes = Left(RemoveQuotes, Len(RemoveQuotes) - 1)

End Function

' return random integer between given lower and upper ranges
Public Function rndint(ByVal lowerbound As Long, ByVal upperbound As Long) As Long
    rndint = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
End Function

' StripNonPrint: Strip non-printable characters and those not supported by Windows from a string
' Trust:
' Arguments: strToStrip
' Returns: string - stripped string
Public Function StripNonPrint(strToStrip As String) As String

    Dim nLengthToStrip As Long, strCharToCheck As String, nAscToCheck As Integer
    Dim lPlace As Long
    
    nLengthToStrip = Len(strToStrip)
    
    ' for each character in the string to strip, only add it to the output string if it meets the criteria
    For lPlace = 1 To nLengthToStrip
        strCharToCheck = Mid$(strToStrip, lPlace, 1)
        nAscToCheck = Asc(strCharToCheck)
        
        ' anything before &H20 is non-printing.  The others excluded are codes Windows does not use
        If (nAscToCheck >= 32 And nAscToCheck <= 126) Or (nAscToCheck >= 145 And nAscToCheck <= 146) Or nAscToCheck >= 160 Then
            StripNonPrint = StripNonPrint & strCharToCheck
        End If
    Next

End Function

' TruncLeft: Truncates the left hand side of the given string with the characters in the regexp until it meets a char not in the regexp
' Trust:
' Arguments: byref strToTrunc, byref strRegExp
' Returns: NONE - operates directly on strToTrunc
Public Sub TruncLeft(ByRef strToTrunc As String, ByRef strRegExp As String)

    ' Sanity - if an empty string has been passed then exit gracefully
    If strToTrunc = "" Then Exit Sub
    
    Dim i As Long: i = 1
    Dim lStrLen As Long: lStrLen = Len(strToTrunc)
    
    ' Find the first character not in the reg exp
    Do While Mid$(strToTrunc, i, 1) Like strRegExp
    
        i = i + 1
        
        ' Don't look any further if we've reached the last character in the string
        If i > lStrLen Then Exit Do
    
    Loop
    
    ' Truncate at this point
    strToTrunc = Right$(strToTrunc, lStrLen - (i - 1))
    
End Sub

' TruncLeft: Truncates the Left hand side of the given string with the characters in the regexp until it meets a char not in the regexp
' Trust:
' Arguments: byref strToTrunc, byref strRegExp
' Returns: NONE - operates directly on strToTrunc
Public Sub TruncRight(ByRef strToTrunc As String, ByRef strRegExp As String)

    ' Sanity - if an empty string has been passed then exit gracefully
    If strToTrunc = "" Then Exit Sub
    
    Dim i As Long: i = 1
    Dim lStrLen As Long: lStrLen = Len(strToTrunc)
    
    ' Find the first character not in the reg exp
    Do While Mid$(strToTrunc, lStrLen - (i - 1), 1) Like strRegExp
    
        i = i + 1
        
        ' Don't look any further if we've reached the last character in the string
        If i > lStrLen Then Exit Do
    
    Loop
    
    ' Truncate at this point
    strToTrunc = Left$(strToTrunc, lStrLen - (i - 1))
    
End Sub
