VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

' Word:  Store the details of a word node and the associated links collection
' Requires (incomplete list): clsLog
' Global vars - gnVersion, gobjLog
' *** Really, Word should be a friend class of Actor
' *** Need to find out how to do this

' *** clsWord requires basGeneral module ***

Option Explicit

Private mlNoNo As Long               ' Unique Word object number.
Private mstrWord As String           ' DEFAULT PROPERTY, actual textual word of the word object
Private mobjParentVocab As clsVocab  ' Vocab to which this word belongs.  Each instance can only belong to one vocab
Private mcolPrevLinks As clsLinks    ' Storage for previous word links
Private mcolNextLinks As clsLinks    ' Storage for next word links
Private mcolConLinks As clsLinks     ' Storage for contextual links

' AddConLink: Add a contextual link to the word
' Trust: HIGH - all node numbers entering the function are valid
' Arguments: lLinkNo - word to link, lConOccs - no of link occurences
' Returns: clsLink - link created
Public Function AddConLink(ByVal lLinkNo As Long, ByVal lConOccs As Long) As clsLink
    
    ' Setup link
    Set AddConLink = mcolConLinks.Add(lLinkNo, lConOccs)
    
    ' Update the vocab list view and vocab contextual count
    With mobjParentVocab
        .lConOccsCount = .lConOccsCount + lConOccs
    End With
    frmVocab.Update Me
    
End Function

' AddCtrlNextLink: Given a word to next link with, add a ctrl word link (for use by ctrl words)
' Trust: LOW
' Arguments lNextWordObjNo
' Returns: link created
Public Function AddCtrlNextLink(ByVal lNextWordObjNo As Long) As clsLink
    
    ' Add the link
    Set AddCtrlNextLink = mcolNextLinks.Add(lNextWordObjNo, 1)

    ' Update the vocab list view to reflect new occurence
    frmVocab.Update Me
    
End Function

' AddCtrlPrevLink: Given a word to previous link with, add a ctrl word link (for use by ctrl words)
' Trust: LOW
' Arguments lPrevWordObjNo
' Returns: link created
Public Function AddCtrlPrevLink(ByVal lPrevWordObjNo As Long) As clsLink
    
    ' Add the link
    Set AddCtrlPrevLink = mcolPrevLinks.Add(lPrevWordObjNo, 1)

    ' Update the vocab list view to reflect new occurence
    frmVocab.Update Me
    
End Function

' AddOrdNextLink: Given a next word, add the appropriate link and perform updates
' Trust: LOW
' Arguments: lNextWordObjNo
' Returns: link created
Public Function AddOrdNextLink(ByVal lNextWordObjNo As Long) As clsLink

    ' Add the next word link
    Set AddOrdNextLink = mcolNextLinks.Add(lNextWordObjNo, 1)

    ' Update the vocab list view and vocab count to reflect new occurence
    With mobjParentVocab
        .lOrdOccsCount = .lOrdOccsCount + 1
    End With
    frmVocab.Update Me
    
End Function

' AddOrdPrevLink: Given a previous word, add the appropriate link and perform updates
' Trust: LOW
' Arguments: lPrevWordObjNo
' Returns: link created
Public Function AddOrdPrevLink(ByVal lPrevWordObjNo As Long) As clsLink
    
    ' Add the next word link
    Set AddOrdPrevLink = mcolPrevLinks.Add(lPrevWordObjNo, 1)

    ' Update the vocab list view and vocab count to reflect new occurence
    With mobjParentVocab
        .lOrdOccsCount = .lOrdOccsCount + 1
    End With
    frmVocab.Update Me
    
End Function

' *** MOTHBALLED, replaced by Len(ConWord()) ***
' ConLen: Return the 'contextual' length of this word.
' Trust: NONE
' Arguments: NONE
' Returns: long - contextual length of word
'Public Function ConLen() As Long
'
'    Dim lConLength As Long, lPlace As Long
'
'    ' count all the alphanumeric characters
'    ConLen = 0
'    For lPlace = 1 To Len(mstrWord)
'
'        If Mid$(mstrWord, lPlace, 1) Like "[0-9A-Za-z]" Then ConLen = ConLen + 1
'
'    Next
'
'End Function

' ConWord: Returns the 'contextual' word - that is, the word stripped of leading and trailing non-alphanumeric chars
' Trust:
' Arguments: NONE
' Returns: string
Public Function ConWord() As String

    ConWord = mstrWord
    
    ' Chop off leading and trailing chars that aren't alphanumeric
    TruncLeft ConWord, "[!A-Za-z0-9]"
    TruncRight ConWord, "[!A-Za-z0-9]"

End Function

' Create: Create the initial data for this word
' Arguments: strWord - text of new word, lObjNo - object number
' Arguments: objParentVocab - vocabulary object to which this word belongs
' Returns: reference to new word
Public Function Create(strWord As String, lObjNo As Long, objParentVocab As clsVocab) As clsWord
    
    ' Make sure the parameters creation is fed are valid, to save trouble later on
    Debug.Assert strWord <> "" And Not objParentVocab Is Nothing
        
    ' Set up parameters
    mstrWord = strWord
    mlNoNo = lObjNo
    Set mobjParentVocab = objParentVocab
    
    Set Create = Me

End Function

' GetRandomLink: returns a Random forward (backward) Link on the selected Word, weighting for number of occurences
' Trust: LOW, only accepts backward and forward directions
' Arguments: direction - indicates whether the link required is a backward or forward link
' Returns: Random link or Nothing if no valid links for this direction exist.
Public Function GetRandomLink(direction As order) As clsWord
    Dim objRandomLink As clsLink
    
    If direction = forward Then
    
        Set objRandomLink = mcolNextLinks.GetRandom
        ' If there are no links to choose at random, return appropriately
        If objRandomLink Is Nothing Then Set GetRandomLink = Nothing: Exit Function
        Set GetRandomLink = mobjParentVocab.Resolve(objRandomLink.lLinkNo)
        Exit Function
        
    ElseIf direction = backward Then
    
        Set objRandomLink = mcolPrevLinks.GetRandom
        ' If there are no links to choose at random, return appropriately
        If objRandomLink Is Nothing Then Set GetRandomLink = Nothing: Exit Function
        Set GetRandomLink = mobjParentVocab.Resolve(objRandomLink.lLinkNo)
        Exit Function
        
    End If
    
    ' Sanity - Shouldn't ever reach this point
    MsgBox "Word:GetRandomLink has been given a direction which isn't forward or backward!"
    Stop
    
End Function

' IsValidTerminator: Assess whether this word has a high enough EOS or SOS weight to be regarded as a terminator
' Arguments: eDirection - whether backward or forward links to be tested, lWeightReq - threshold to be reached or exceeded
' Trust:
' Returns: bool
Public Function IsValidTerminator(eDirection As order, fProportionReq As Single) As Boolean

    Dim objCtrlLink As clsLink
    
    ' Sanity - con should never be given as the direction
    If eDirection = con Then gobjLog.Log mstrWord + ":IsValidTerminator called with con mistakenly as order!", Warn: IsValidTerminator = False: Exit Function
    
    ' Execute slightly different code depending on whether we're testing forward or backward links
    If eDirection = forward Then
    
        ' If the word has no links (perhaps because other words have been deleted) then assume it can terminate the sentence
        If colNextLinks.Count <= 0 Then
            IsValidTerminator = True: Exit Function
       
        Else
            ' If the word has a EOS link
            Set objCtrlLink = colNextLinks.Find(conEOSObjNo)
                If Not objCtrlLink Is Nothing Then
                    
                    ' See if the EOS link is weighted heavily enough for it to be regarded as a valid terminator
                    If objCtrlLink.lOccs / colNextLinks.lOccs >= fProportionReq Then
                    
                        IsValidTerminator = True: Exit Function
                        
                    End If
                    
                End If
        End If
        
    ' Direction is backward
    Else
    
        ' If the word has no links (perhaps because other words have been deleted) then assume it can terminate the sentence
        If colPrevLinks.Count <= 0 Then
            IsValidTerminator = True: Exit Function
       
        Else
            ' If the word has a SOS link
            Set objCtrlLink = colPrevLinks.Find(conSOSObjNo)
                If Not objCtrlLink Is Nothing Then
                    
                    ' See if the EOS link is weighted heavily enough for it to be regarded as a valid terminator
                    If objCtrlLink.lOccs / colPrevLinks.lOccs >= fProportionReq Then

                        IsValidTerminator = True: Exit Function
                        
                    End If
                    
                End If
        End If
        
    End If
    
    IsValidTerminator = False
            
End Function

' Load: Load the contents of the Word specified by fileno
' Arguments: fileno - File number being used for loading
' Arugments: objParentVocab - clsVocab to which this word belongs
' Returns: non-zero if error
Public Function Load(ByVal fileno As Integer, ByVal objParentVocab As clsVocab) As Long

    On Error GoTo ErrorHandler
    
    Dim nWordLength As Integer, lPrevLinksCount As Long, lNextLinksCount As Long, lConLinksCount As Long
    Dim strLoadedWord As String, lObjNo As Long
    Dim yCheck As Byte
    Dim lReturnCode&
    
    Get fileno, , nWordLength                   ' Need to retrieve length of word before retrieving word
    strLoadedWord = Input(nWordLength, fileno)  ' Retrieve string and object number to temporary variables, so we can invoke Create
    Get fileno, , lObjNo
    Create strLoadedWord, lObjNo, objParentVocab
    
    ' Debug.Print "Word.Load(): "; mstrWord  ' ***DEBUG LINE
    
    ' Load other properties which aren't setup by Create
    Get fileno, , lPrevLinksCount
    Get fileno, , lNextLinksCount
    Get fileno, , lConLinksCount
    
    ' Check integrity of file
    If gnVersion >= 20 Then
        Get fileno, , yCheck
        If yCheck <> MAGICBYTE Then
            gobjLog.Log "Load file loses integrity after loading variables in obj: " + CStr(lObjNo) + " - " + strLoadedWord, Error
            Err.Raise errMagicMisplaced
        End If
    End If
    
    ' Link loading
    lReturnCode = mcolPrevLinks.Load(fileno, lPrevLinksCount): If lReturnCode Then Err.Raise lReturnCode
    
    ' Check integrity of file
    If gnVersion >= 20 Then
        Get fileno, , yCheck
        If yCheck <> MAGICBYTE Then
            gobjLog.Log "Load file loses integrity loading prev links in obj: " + CStr(lObjNo) + " - " + strLoadedWord, Error
            Err.Raise errMagicMisplaced
        End If
        
    End If
    
    lReturnCode = mcolNextLinks.Load(fileno, lNextLinksCount): If lReturnCode Then Err.Raise lReturnCode
    
    ' Check integrity of file
    If gnVersion >= 20 Then
        Get fileno, , yCheck
        If yCheck <> MAGICBYTE Then
            gobjLog.Log "Load file loses integrity loading next links in obj: " + CStr(lObjNo) + " - " + strLoadedWord, Error
            Err.Raise errMagicMisplaced
        End If
    End If
    
    lReturnCode = mcolConLinks.Load(fileno, lConLinksCount): If lReturnCode Then Err.Raise lReturnCode
    
    ' Check integrity of file
    If gnVersion >= 20 Then
        Get fileno, , yCheck
        If yCheck <> MAGICBYTE Then
            gobjLog.Log "Load file loses integrity loading con links in obj: " + CStr(lObjNo) + " - " + strLoadedWord, Error
            Err.Raise errMagicMisplaced
        End If
    End If
    
    ' Signal no error
    Load = 0
    
    Exit Function
    
ErrorHandler:

    Load = Err.Number
    
End Function

' RmAllConLinks: Remove all contextual links
' Trust:
' Arguments: NONE
' Returns: clsWords - collection of words which have been affected by the removal of these links
Public Function RmAllConLinks() As clsWords

    Set RmAllConLinks = New clsWords
    
    Dim objLink As clsLink, objAffectedWord As clsWord
    Dim lLinksRemoved As Long
    
    ' For all the words to which this word contextually links, remove their contextual links to this word.
    For Each objLink In mcolConLinks.mcolLinks
    
        Set objAffectedWord = mobjParentVocab.Resolve(objLink.lLinkNo)
        
        ' Sanity - if the object number has not resolved something has gone wrong previously.  Write error but continue
        If objAffectedWord Is Nothing Then
            gobjLog.Log "clsWord:RmAllConLinks tried to remove objno " + CStr(objLink.lLinkNo) + " which no longer exists!", Warn
            Exit For
        End If
        
        ' Remove the link and update the internal record of link instances
        lLinksRemoved = objAffectedWord.colConLinks.Remove(mlNoNo)
        
        ' Update the internal record of link instances if the affected word is not a control word
        If IsRealWord(objAffectedWord.nNoNo) Then mobjParentVocab.lConOccsCount = mobjParentVocab.lConOccsCount - lLinksRemoved
        
        ' Add to words that need to be updated
        RmAllConLinks.Add objAffectedWord, False
        
    Next
    
    ' Update the vocab list view and vocab contextual count
    With mobjParentVocab
        .lConOccsCount = .lConOccsCount - mcolConLinks.Count
    End With
    frmVocab.Update Me

    ' Forget about the old collection of contextual links and create a new empty one
    Set mcolConLinks = New clsLinks
    
End Function

' RmAllNextLinks: Remove all contextual links
' Trust:
' Arguments: NONE
' Returns: clsWords - collection of words which have been affected by the removal of these links
Public Function RmAllNextLinks() As clsWords

    Set RmAllNextLinks = New clsWords
    
    Dim objLink As clsLink, objAffectedWord As clsWord
    Dim lLinksRemoved As Long
    
    ' For all the words to which this word contextually links, remove their contextual links to this word.
    For Each objLink In mcolNextLinks.mcolLinks
    
        Set objAffectedWord = mobjParentVocab.Resolve(objLink.lLinkNo)
        
        ' Sanity - if the object number has not resolved something has gone wrong previously.  Write error but continue
        If objAffectedWord Is Nothing Then
            gobjLog.Log "clsWord:RmAllNextLinks tried to remove objno " + CStr(objLink.lLinkNo) + " which no longer exists!", Warn
            Exit For
        End If
        
        ' Remove the link and update the internal record of link instances
        lLinksRemoved = objAffectedWord.colNextLinks.Remove(mlNoNo)
        
        ' Update the internal record of link instances if the affected word is not a control word
        If IsRealWord(objAffectedWord.nNoNo) Then mobjParentVocab.lOrdOccsCount = mobjParentVocab.lOrdOccsCount - lLinksRemoved
        
        ' Add to words that need to be updated
        RmAllNextLinks.Add objAffectedWord, False
        
    Next
    
    ' Update the vocab list view and vocab contextual count
    With mobjParentVocab
        .lOrdOccsCount = .lOrdOccsCount - mcolNextLinks.Count
    End With
    frmVocab.Update Me

    ' Forget about the old collection of contextual links and create a new empty one
    Set mcolNextLinks = New clsLinks
    
End Function

' RmAllPrevLinks: Remove all contextual links
' Trust:
' Arguments: NONE
' Returns: clsWords - collection of words which have been affected by the removal of these links
Public Function RmAllPrevLinks() As clsWords

    Set RmAllPrevLinks = New clsWords
    
    Dim objLink As clsLink, objAffectedWord As clsWord
    Dim lLinksRemoved As Long
    
    ' For all the words to which this word contextually links, remove their contextual links to this word.
    For Each objLink In mcolPrevLinks.mcolLinks
    
        Set objAffectedWord = mobjParentVocab.Resolve(objLink.lLinkNo)
        
        ' Sanity - if the object number has not resolved something has gone wrong previously.  Write error but continue
        If objAffectedWord Is Nothing Then
            gobjLog.Log "clsWord:RmAllPrevLinks tried to remove objno " + CStr(objLink.lLinkNo) + " which no longer exists!", Warn
            Exit For
        End If
        
        ' Remove the link and update the internal record of link instances
        lLinksRemoved = objAffectedWord.colPrevLinks.Remove(mlNoNo)
        
        ' Update the internal record of link instances if the affected word is not a control word
        If IsRealWord(objAffectedWord.nNoNo) Then mobjParentVocab.lOrdOccsCount = mobjParentVocab.lOrdOccsCount - lLinksRemoved
        
        ' Add to words that need to be updated
        RmAllPrevLinks.Add objAffectedWord, False
        
    Next
    
    ' Update the vocab list view and vocab contextual count
    With mobjParentVocab
        .lOrdOccsCount = .lOrdOccsCount - mcolPrevLinks.Count
    End With
    frmVocab.Update Me

    ' Forget about the old collection of contextual links and create a new empty one
    Set mcolPrevLinks = New clsLinks
    
End Function

' Save: Save the contents of this Word to the file specified in filename
' Trust:
' Arguments: fileno
' Returns: non-zero if error
Public Function Save(ByRef fileno As Integer) As Long

    On Error GoTo ErrorHandler

    ' Put Word saving code here
    Put fileno, , CInt(Len(mstrWord))  ' Need to write length of string, since in Binary mode only the string itself is saved
    Put fileno, , mstrWord
    Put fileno, , mlNoNo
    Put fileno, , mcolPrevLinks.Count()
    Put fileno, , mcolNextLinks.Count()
    Put fileno, , mcolConLinks.Count()
    
    ' File checking
    Put fileno, , MAGICBYTE
    
    ' Save the links
    Save = mcolPrevLinks.Save(fileno): If Save Then gobjLog.Log "Save failed for prevlinks in " + mstrWord, Gen: Exit Function
    Put fileno, , MAGICBYTE
    Save = mcolNextLinks.Save(fileno): If Save Then gobjLog.Log "Save failed for nextlinks in " + mstrWord, Gen: Exit Function
    Put fileno, , MAGICBYTE
    Save = mcolConLinks.Save(fileno): If Save Then gobjLog.Log "Save failed for contextual links in " + mstrWord, Gen: Exit Function
    Put fileno, , MAGICBYTE
    
    ' Debug.Print "Saved: "; mstrWord  ' ***DEBUG LINE
    
    ' Signal everything went okay
    Save = 0
    
    Exit Function
    
ErrorHandler:
    Save = Err.Number
    
End Function

' TearDown: Releases all references to parent objects and other words contained in the word so that it can terminate
' Trust: N/A
' Arguments: NONE
' Returns: NONE
Public Sub TearDown()

    Set mobjParentVocab = Nothing
    
End Sub

' Read-write properties
Public Property Get strWord() As String: strWord = mstrWord: End Property
Public Property Let strWord(strWord As String): mstrWord = strWord: End Property

' Read only properties
' Property Get colConLinks: returns the contextual links collection for this word
Public Property Get colConLinks() As clsLinks: Set colConLinks = mcolConLinks: End Property
' Property Get colPrevLinks: return the previous links collection for this word
Public Property Get colPrevLinks() As clsLinks: Set colPrevLinks = mcolPrevLinks: End Property
' Property Get colPrevLinks: return the next links collection for this word
Public Property Get colNextLinks() As clsLinks: Set colNextLinks = mcolNextLinks: End Property
Public Property Get nNoNo() As Long: nNoNo = mlNoNo: End Property
Public Property Get objParentVocab() As clsVocab: Set objParentVocab = mobjParentVocab: End Property

' Debug routine to show the links of a Word in the Immediate window
'Public Sub DebugPrintLinks()
'    Dim i As Integer: i = 0
'    Dim text As String
'    Do While i < mcolLinks.Count()
'        ' Need to detect if variant is an object because if so, the test against EOS (an integer) throws an error
'        If Not mpLinks(i).linextno = NONODE Then
'            text = mobjParentVocab.Resolve(mpLinks(i).linextno).mstrWord
'        Else
'            text = "EOS"
'        End If
'        ' Debug.Print "<link"; i; "for "; word; " to (#"; Format(mpLinks(i).linextno); ") "; text; " - Occs: "; Format(mpLinks(i).lioccs)
'        i = i + 1
'    Loop
'End Sub

Private Sub Class_Initialize()
    ' Initialize variables
    mstrWord = "ERROR"
    Set mobjParentVocab = Nothing
    
    Set mcolPrevLinks = New clsLinks
    Set mcolNextLinks = New clsLinks
    Set mcolConLinks = New clsLinks
    
End Sub

Private Sub Class_Terminate()
    ' Debug.Print "*** "; mstrWord; " deleted ***"
End Sub
