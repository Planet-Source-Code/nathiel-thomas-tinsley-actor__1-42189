Attribute VB_Name = "basMainFunctions"
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

Option Explicit

' Maxnumber of characters allowed in a sentence.  Same as upper limit of integer type
Private Const MAXCHARS = 32767

' AnalyzeSentence: Cut the next sentence from a string and tokenise
' Trust: LOW, should be able to cope with any string, even nonsense stuff
' Arguments: s - string to analyze
' Returns: variant - the reference to the array of tokens
Public Function AnalyzeSentence(ByRef strUserInput As String) As Variant
    Dim strUserSentence As String
    
    ' Cut the next sentence off
    strUserSentence = SplitOffSentence(strUserInput)

    ' Tokenise and return
    AnalyzeSentence = Tokenise(strUserSentence)
    ' PrintTokens word   ' ***DEBUG LINE
    
End Function

' DeGenWordKey: Given a key, return the word it describes
' Trust: MEDIUM, effectively checks the string given to be a key
' Arguments: strKey
' Returns: word - word associated with the given key
Public Function DeGenWordKey(ByVal strKey As String) As clsWord

    ' Return the associated word
    Set DeGenWordKey = gobjCurrActor.colVocab.Resolve(KeyToObjNo(strKey))
    
End Function

' GenWordKey: Given a word, return the appropriate key
' GenWordKey: We have to stick a '#' in the front of Word number to generate the key
' GenWordKey: because the ListView control doesn't like keys starting with digits...)
' Trust: MEDIUM, checks if an object has been passed but has to assume this is a word object
' Arguments: objWordToKey
' Returns: string
Public Function GenWordKey(objWordToKey As Variant) As String

    ' Distinguish between a passed object and a passed number, and act accordingly
    Select Case VarType(objWordToKey)
    
        Case vbObject
            ' If an object is passed then assume it's a word (weak, but don't know an elegant solution in VB5)
            ' Sanity check - don't continue if the word reference passed is empty
            If objWordToKey Is Nothing Then GenWordKey = "": Exit Function
            
            ' Setup word component of key
            GenWordKey = CStr(objWordToKey.nNoNo)
            
        Case vbLong
        
            ' Convert to string
            GenWordKey = CStr(objWordToKey)
            
        Case Else
        
            ' Sanity check - Don't deal with anything apart from objects and longs
            Stop
            
    End Select
    
    ' Construct the listview key and return it
    GenWordKey = "#" + GenWordKey
    
End Function

' IsRealWord: For any object number, indicate whether it's a real word or a control word
' Trust: LOW
' Arguments: lWordToCheck
' Returns: bool
Public Function IsRealWord(ByVal lWordToCheck As Long) As Boolean

    If lWordToCheck >= conFirstWordObjNo Then IsRealWord = True Else IsRealWord = False
    
End Function

' IsValidKey: For any string, indicate whether it's in the prescribed format
' Trust: LOW
' Arguments: strKeyToCheck
' Returns: bool
Public Function IsValidKey(ByVal strKeyToCheck As String) As Boolean

    IsValidKey = True
    ' Check the first character is a #
    If Left$(strKeyToCheck, 1) <> "#" Then IsValidKey = False
    ' Check the remainder of the string is alphanumeric
    If Not IsNumeric(Mid$(strKeyToCheck, 2)) Then IsValidKey = False
    
End Function

' KeyToObjNo: Convert a given key into an object number
' Trust: LOW, checks conversion results in a numeric key
' Arguments: strKeyToTransform
' Returns: long - object number
Public Function KeyToObjNo(strKeyToTransform As String) As Long

    Dim strTransformedKey As String

    ' Transform the key into the word object number
    strTransformedKey = Mid$(strKeyToTransform, 2)
    
    ' Sanity - Check this is a number
    If Not IsNumeric(strTransformedKey) Then MsgBox "MainFunction:KeyToObjNo called with " + strKeyToTransform + " transformed to " + strTransformedKey + " which is invalid!", vbExclamation: Stop
    
    ' Return object number
    KeyToObjNo = CLng(strTransformedKey)

End Function

' Reporting: Check if the set level of reporting matches the enumerated argument level and return true or false
' Trust: N/A
' Arguments: eRepLevel
' Returns: Boolean
Public Function Reporting(ByVal eRepLevel As geRepLevel) As Boolean

    If gobjOptions.meRepLevel = eRepLevel Then Reporting = True Else Reporting = False
    
End Function

' SplitOffSentence: Return the next sentence from the string s and split this off from s
' SplitOffSentence: Need to handle situation when user enters CTRL+ENTER in user textbox... it causes the textbox to pass a CRLF!
' Arguments: sentence - string to split
' Returns: string - split off sentence
Private Function SplitOffSentence(ByRef sentence As String) As String
    Dim i As Integer, l As Integer: i = 1: l = Len(sentence)
    Dim c As String
    
    ' If there's no sentence left, then pop the function
    If l = 0 Then SplitOffSentence = "": Exit Function
    
    c = Mid$(sentence, i, 1)
    ' Snip off the sentence at the next sentence ender or if it gets too big (unlikely)
    ' or if we reach the end of the argument string altogether
    Do While c <> "." And c <> "?" And c <> "!" And c <> vbCr And c <> vbLf And i < MAXCHARS And i <> l
        i = i + 1
        c = Mid$(sentence, i, 1)
    Loop
    
    ' If we've come to a halt because of a sentence ender
    ' Then continue getting the sentence until there are no more sentence enders (to handle more than one !, etc)
    If i < MAXCHARS And i <> l Then
        c = Mid$(sentence, i + 1, 1)
        Do While (c = "." Or c = "?" Or c = "!" Or c = vbCr Or c = vbLf)
            i = i + 1
            ' Pop sentence if we've reached the end and the mid statement would be illegal
            If i = l Or i = MAXCHARS Then Exit Do Else c = Mid$(sentence, i + 1, 1)
        Loop
    End If
    
    ' Return this sentence and split it off from the original string, adjusting the i counter to reflect extra i counted
    SplitOffSentence = Left$(sentence, i)
    sentence = Right$(sentence, l - i)

End Function

' Tokenise: Process sentence string and return a reference to a created array of strings
' Tokenise: Need to reverse word order so that links between words can be set up later
' Arguments: Original user sentence string
' Returns: returns 0 if failed, else array of strings of tokens backwards as they appear in the input sentence
Private Function Tokenise(User As String) As Variant

    Dim word() As String
    Dim pos As Integer, i As Integer, j As Integer
    
    pos = Len(User$): i = 0
    
    ' Work through the given user string backwards
    Do While pos > 0
    
        ' Step pos through spaces until the next word is reached, stopping if necessary
        Dim c As String: c = Mid$(User, pos, 1)
        Do Until c <> " " And c <> "." And c <> vbCr And c <> vbLf
            pos = pos - 1
            If pos <= 0 Then Exit Do
            c = Mid$(User, pos, 1)
        Loop
        If pos <= 0 Then Exit Do  ' exit loop if we reach the end of the sentence to prevent a fault with mid
        
        ReDim Preserve word(i + 1)  ' Enlarge the dynamic array by one word
        
        ' Get the size and position of the word immediately before the pos pointer
        j = 0
        ' Ignore full stops as well as spaces
        ' XXX Might be a better write possible using instr XXX
        c = Mid$(User, pos, 1)
        Do While c <> " " And c <> "." And c <> vbCr And c <> vbLf
            pos = pos - 1: j = j + 1: If pos <= 0 Then Exit Do
            c = Mid$(User, pos, 1)
        Loop
        
        ' Get the token
        word$(i) = Mid$(User$, pos + 1, j)
        
        ' Move onto the next word in the word array of strings
        i = i + 1
        
    Loop
    
    ' Convert the last token to lower case if appropriate
    ' Appropriate condition is if any words have been parsed and if the word isn't "I"
    If i > 0 Then
        If word(i - 1) <> "I" Then word(i - 1) = StrConv(word(i - 1), vbLowerCase)
    End If
    
    ' Add an empty string as the last word to signal the end in the array of strings
    ReDim Preserve word(i + 1)
    word$(i) = ""
    Tokenise = word  ' I think the local array word is now preserved, since Tokenise is a reference
    
End Function
