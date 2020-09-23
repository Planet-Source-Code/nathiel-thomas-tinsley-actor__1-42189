VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVocab"
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

' clsVocab: Aim is to handle all vocabulary related functions based on the words collection class
' Global vars - gnVersion

Option Explicit

Private mobjOwnerActor As clsActor     ' Actor to which this vocabulary object belongs
Private mlNextObjNo As Long            ' Object no to use for the next new word
Private mlWordsCount As Long           ' Number of words
Private mlOrdOccsCount As Long         ' Number of word occurences, as opposed to number of words
Private mlConOccsCount As Long         ' Number of contextual link occurences
Private mcolControlWords As clsWords   ' The collection of word objects which are control 'words' (eg EOS)
Private mcolVocabWords As clsWords
Public mcolLookUp As clsLookUp         ' Contains information linking object numbers to words

' Add: If it doesn't already exist, create a word from the string and add it to the collection
' Trust:
' Arguments: word - string to create a word
' Returns: word - word object created
Public Function Add(word As String) As clsWord

    ' Debug.Print "START clsVocab:Add"
    
    Dim objNode As clsWord, nReturn As Integer
    
    ' Check collection to see if we already have this word.
    Set objNode = mcolVocabWords.Find(word)
    
    ' If we don't have it, then create it and add it to the collection
    If objNode Is Nothing Then
        
        Set objNode = New clsWord
        mcolVocabWords.Add objNode.Create(word, NextObjNo(), Me)
        
        ' Add new word to lookup table
        nReturn = mcolLookUp.Add(objNode.nNoNo, word)
        If nReturn Then MsgBox "Error " + CStr(nReturn) + " when adding word to lookup!", vbCritical
                 
        ' Update number of words with the new one
        mlWordsCount = mlWordsCount + 1
        
        ' Update the Vocabulary list view when a new word is added XXX GOT TO MOVE THIS XXX
        frmVocab.AddWord objNode
        
        ' Debug.Print "clsVocab:Add added "; objNode.strWord; " ObjNo:"; objNode.nNoNo
    End If
    
    Set Add = objNode
    
    ' Debug.Print "END clsVocab:Add"
End Function

' Create: Initialise actor owner in clsVocab object
' Trust: LOW
' Arguments: objOwnerActor
' Returns: NONE
Public Sub Create(objOwnerActor As clsActor)

    If Not objOwnerActor Is Nothing Then
        Set mobjOwnerActor = objOwnerActor
    Else
        MsgBox "Error: clsVocab:Create called with Nothing!", vbCritical
        Stop
    End If
    
End Sub

' Edit: Correct the string of a word in the vocabulary
' Trust: LOW, checks word actually belongs to the vocabulary collection of this object
' Arguments: objEditWord, strNewString
' Returns: NONE
Public Sub Edit(objEditWord As clsWord, strNewString As String)

    ' Signal that the current actor now differs from the last saved version
    mobjOwnerActor.bChangedFlag = True
    
    ' Check word is in collection and change if it is
    If objEditWord.objParentVocab Is Me Then
        objEditWord.strWord = strNewString
    Else
        MsgBox "clsVocab:Edit was called with a word which was not contained in its vocabulary", vbExclamation
    End If
    
End Sub

' Find: Find word object associated with the search string
' Trust: LOW, strSearch can be anything.
' Arguments: strSearch
' Returns: word - word object if found, Nothing if not
Public Function Find(strSearch As String) As clsWord
    Set Find = colVocabWords.Find(strSearch)
End Function

' GetOrdRnd: Get a random word weighting by ordinary links
' Trust:
' Arguments: NONE
' Returns: clsWord - chosen word
Public Function GetOrdRnd() As clsWord

    Dim strLog$

    ' Sanity - if there are no occurences then we cannot pick a word
    If mlOrdOccsCount = 0 Then Set GetOrdRnd = Nothing: Exit Function

    ' *** If there are words then if clsWords:GetOrdRnd returns nothing we must have supplied an inaccurate occs figure
    ' *** However, this is not fatal, so log an error but keep trying until we do get a word
    Do Until Not GetOrdRnd Is Nothing
    
        Set GetOrdRnd = mcolVocabWords.GetOrdRnd(mlOrdOccsCount)
        
        If GetOrdRnd Is Nothing Then
            strLog = "The main vocabulary collection :GetOrdRnd was passed " + CStr(mlOrdOccsCount) + " which is too many occurences."
            gobjLog.Log strLog, Warn
        End If
        
    Loop
    
End Function

' Load: Load the members for this vocab class from an open file
' Arguments: fileno - file number of vocab file
' Returns: non-zero if error
Public Function Load(ByRef fileno As Integer) As Long

    On Error GoTo ErrorHandler

    Dim objWord As clsWord, i As Integer, nReturn As Integer
    Dim objFoundWord As clsWord
    Dim yCheck As Byte
    Dim lReturnCode&
    
    ' Load variable details
    Get fileno, , mlWordsCount
    Get fileno, , mlOrdOccsCount
    Get fileno, , mlConOccsCount
    Get fileno, , mlNextObjNo
    
    ' Check integrity of file
    If gnVersion >= 20 Then
        Get fileno, , yCheck
        If yCheck <> MAGICBYTE Then
            gobjLog.Log "Load file loses integrity loading clsVocab variables", Error
            Err.Raise errMagicMisplaced
        End If
    End If

    ' Load control word details
    lReturnCode = LoadCtrlWords(fileno): If lReturnCode Then Err.Raise lReturnCode
    
    ' Load words
    For i = 1 To mlWordsCount
    
        ' Create a word and load its details from the file
        Set objWord = New clsWord
        lReturnCode = objWord.Load(fileno, Me): If lReturnCode Then Err.Raise lReturnCode
        
        ' To maintain backwards compatibility with 0.1.7, handle words which would be duplicates if converted to lowercase
        ' e.g. "Aliens" and "aliens"
        Set objFoundWord = mcolVocabWords.mcolWords(objWord.strWord)
        If Not objFoundWord Is Nothing Then
        
            ' Create entry in lookup which points duplicate word object numbers to original word
            nReturn = mcolLookUp.Add(objWord.nNoNo, objFoundWord.strWord)
            If nReturn Then gobjLog.Log "clsCollection:Add returned error " + CStr(nReturn) + " when called by clsVocab:Load for " + CStr(objWord.nNoNo) + " - " + objFoundWord.strWord, Error
            
            ' Adjust the newly loaded statistics to reflect the fact we no longer have the duplicate word in the vocab.
            mlWordsCount = mlWordsCount - 1
            mlOrdOccsCount = mlOrdOccsCount - objWord.colPrevLinks.Count - objWord.colNextLinks.Count
            mlConOccsCount = mlConOccsCount - objWord.colConLinks.Count
            
        Else
            
            ' Add new word to vocabulary display list
            frmVocab.AddWord objWord
            
            ' Add Word to the vocab
            mcolVocabWords.Add objWord
            
            ' Create entry in lookup table
            With objWord
                nReturn = mcolLookUp.Add(.nNoNo, .strWord)
                If nReturn Then gobjLog.Log "clsWord:Load failed with error " + CStr(nReturn) + " calling clsLookUp:Add on " + CStr(.nNoNo) + " - " + .strWord, Error
            End With
            
        End If
        
    Next
    
    ' Signal no error
    Load = 0
    
    Exit Function
    
ErrorHandler:
    
    Load = Err.Number
    
End Function

' Merge: Merge two words
' Arguments: objTargetMergeWord, objSourceMergeWord
' Returns: NONE
Public Sub Merge(objTargetMergeWord As clsWord, objSourceMergeWord As clsWord)

    ' Signal that the current actor now differs from the last saved version
    mobjOwnerActor.bChangedFlag = True
    
    ' Copy all the links to previous words from the source to the target and temporarily update internal statistics
    mlOrdOccsCount = mlOrdOccsCount + objTargetMergeWord.colPrevLinks.Copy(objSourceMergeWord.colPrevLinks)
    
    ' Redirect the links pointing to the source word as a next word to the target word instead
    Dim objLinkToSourceWord As clsLink
    Dim objAffectedWord As clsWord
    For Each objLinkToSourceWord In objSourceMergeWord.colPrevLinks.mcolLinks
        Set objAffectedWord = Resolve(objLinkToSourceWord.lLinkNo)
        objAffectedWord.colNextLinks.Redirect objSourceMergeWord.nNoNo, objTargetMergeWord.nNoNo
    Next
    
    ' Copy all the next word links from the source to the target and temporarily update internal statistics
    mlOrdOccsCount = mlOrdOccsCount + objTargetMergeWord.colNextLinks.Copy(objSourceMergeWord.colNextLinks)
    
    ' Redirect the links pointing to the source word as a previous word to the target word instead
    Set objLinkToSourceWord = Nothing
    Set objAffectedWord = Nothing
    For Each objLinkToSourceWord In objSourceMergeWord.colNextLinks.mcolLinks
        Set objAffectedWord = Resolve(objLinkToSourceWord.lLinkNo)
        objAffectedWord.colPrevLinks.Redirect objSourceMergeWord.nNoNo, objTargetMergeWord.nNoNo
    Next
    
    ' Copy all the contextual links from the source to the target
    objTargetMergeWord.colConLinks.Copy objSourceMergeWord.colConLinks
    
    ' Redirect the links on other words associated with the source word to the new target word instead
    Set objLinkToSourceWord = Nothing
    Set objAffectedWord = Nothing
    For Each objLinkToSourceWord In objSourceMergeWord.colConLinks.mcolLinks
        Set objAffectedWord = Resolve(objLinkToSourceWord.lLinkNo)
        objAffectedWord.colConLinks.Redirect objSourceMergeWord.nNoNo, objTargetMergeWord.nNoNo
    Next
    
    ' Remove the source word from the vocabulary pool
    Remove objSourceMergeWord
     
End Sub

' PurgeConExcluded: Purge contextualisation links from words in the vocabulary which appear in the exclusion list
' Trust:
' Arguments: NONE
' Returns: clsWords - the collection of words actually excluded
Public Function PurgeConExcluded() As clsWords

    Set PurgeConExcluded = New clsWords

    Dim colAffectedWords As clsWords
    Dim objWord As clsWord, objAffectedWord As clsWord
    
    ' Remove all the contextual links from any word which appears in the exclusion list
    For Each objWord In mcolVocabWords.mcolWords
    
        With objWord
    
            ' Look at the 'contextual' word
            If mobjOwnerActor.mcolExclusions(.ConWord()) <> "" Then
            
                Set colAffectedWords = .RmAllConLinks()
            
                ' Add the removed word to the display update list
                PurgeConExcluded.Add objWord, False
    
                ' Add the words affected by removing this word to the display update list
                For Each objAffectedWord In colAffectedWords.mcolWords
                    PurgeConExcluded.Add objAffectedWord, False
                Next
        
            End If
        
        End With
    
    Next
    
End Function

' Remove: Removes the word given by setting redirect.  Also updates word and instance counts
' Remove: default behaviour is to end the sentence
' Arguments: objWordToRemove - reference to word to be removed
' Returns: clsWords - collection of words which need to be updated as a result of the word removal
Public Function Remove(objWordToRemove As clsWord) As clsWords

    Set Remove = New clsWords
    Dim objWord As clsWord
    Dim colWordsAffected As clsWords
    
    ' Signal that the current actor now differs from the last saved version
    mobjOwnerActor.bChangedFlag = True
    
    ' Remove all the links from the word, ordinary and contextual
    Set colWordsAffected = objWordToRemove.RmAllNextLinks()
    For Each objWord In colWordsAffected.mcolWords
        Remove.Add objWord, False
    Next
    
    Set colWordsAffected = objWordToRemove.RmAllPrevLinks()
    For Each objWord In colWordsAffected.mcolWords
        Remove.Add objWord, False
    Next
    
    Set colWordsAffected = objWordToRemove.RmAllConLinks()
    For Each objWord In colWordsAffected.mcolWords
        Remove.Add objWord, False
    Next
    
'    ' For all the words to which this is previously linked, remove their links to this word
'    Dim objLinkToAffectedWord As clsLink
'    Dim objAffectedWord As clsWord
'    Dim lLinksRemoved As Long
'    For Each objLinkToAffectedWord In objWordToRemove.colPrevLinks.mcolLinks
'
'        Set objAffectedWord = Resolve(objLinkToAffectedWord.lLinkNo)
'
'        ' Sanity - if the object number has not resolved something has gone wrong previously.  Write error but continue
'        If objAffectedWord Is Nothing Then
'            Debug.Print "clsVocab:Remove tried to remove objno " + CStr(objLinkToAffectedWord.lLinkNo) + " which no longer exists!"
'            Exit For
'        End If
'
'        ' Remove the link
'        lLinksRemoved = objAffectedWord.colNextLinks.Remove(objWordToRemove.nNoNo)
'
'        ' Update the internal record of link instances if the affected word is not a control word
'        If IsRealWord(objAffectedWord.nNoNo) Then mlOrdOccsCount = mlOrdOccsCount - lLinksRemoved
'
'        ' Add to words that need to be updated
'        colWordsAffected.Add objAffectedWord, False
'
'    Next
'
'    ' For all the words to which this word links next, remove their links to this word (as a previous word)
'    For Each objLinkToAffectedWord In objWordToRemove.colNextLinks.mcolLinks
'
'        Set objAffectedWord = Resolve(objLinkToAffectedWord.lLinkNo)
'
'        ' Sanity - if the object number has not resolved something has gone wrong previously.  Write error but continue
'        If objAffectedWord Is Nothing Then
'            Debug.Print "clsVocab:Remove tried to remove objno " + CStr(objLinkToAffectedWord.lLinkNo) + " which no longer exists!"
'            Exit For
'        End If
'
'        ' Remove the link and update the internal record of link instances
'        lLinksRemoved = objAffectedWord.colPrevLinks.Remove(objWordToRemove.nNoNo)
'
'        ' Update the internal record of link instances if the affected word is not a control word
'        If IsRealWord(objAffectedWord.nNoNo) Then mlOrdOccsCount = mlOrdOccsCount - lLinksRemoved
'
'        ' Add to words that need to be updated
'        colWordsAffected.Add objAffectedWord, False
'
'    Next
'
'    ' For all the words to which this word contextually links, remove their contextual links to this word.
'    For Each objLinkToAffectedWord In objWordToRemove.colConLinks.mcolLinks
'
'        Set objAffectedWord = Resolve(objLinkToAffectedWord.lLinkNo)
'
'        ' Sanity - if the object number has not resolved something has gone wrong previously.  Write error but continue
'        If objAffectedWord Is Nothing Then
'            Debug.Print "clsVocab:Remove tried to remove objno " + CStr(objLinkToAffectedWord.lLinkNo) + " which no longer exists!"
'            Exit For
'        End If
'
'        ' Remove the link and update the internal record of link instances
'        lLinksRemoved = objAffectedWord.colConLinks.Remove(objWordToRemove.nNoNo)
'
'        ' Update the internal record of link instances if the affected word is not a control word
'        If IsRealWord(objAffectedWord.nNoNo) Then mlConOccsCount = mlConOccsCount - lLinksRemoved
'
'        ' Add to words that need to be updated
'        colWordsAffected.Add objAffectedWord, False
'
'    Next
'
'    ' Now update the internal record of link instances by the links being directly removed on the deleted word
'    Dim objLinkToRemove As clsLink
'
'    For Each objLinkToRemove In objWordToRemove.colNextLinks.mcolLinks
'        mlOrdOccsCount = mlOrdOccsCount - objLinkToRemove.lOccs
'    Next
'
'    For Each objLinkToRemove In objWordToRemove.colPrevLinks.mcolLinks
'        mlOrdOccsCount = mlOrdOccsCount - objLinkToRemove.lOccs
'    Next
'
'    For Each objLinkToRemove In objWordToRemove.colConLinks.mcolLinks
'        mlConOccsCount = mlConOccsCount - objLinkToRemove.lOccs
'    Next
        
'    ' Update the total 'weight' (occurences) of all words by deleting the influence of the word to be
'    '   deleted, omitting control words
'    Dim objLinkToRemove As clsLink
'    For Each objLinkToRemove In objWordToRemove.colNextLinks
'        With objLinkToRemove
'            ' For a link to an ordinary word, update the total weight to remove both weightings
'            ' ### Needs to be twice because we're counting both backward and forward links
'            ' ### If the link is to a control word, then only 1 weight needs to be removed
'            If .lLinkNo >= conFirstWordObjNo Then
'                mlOrdOccsCount = mlOrdOccsCount - .lOccs * 2
'            Else
'                mlOrdOccsCount = mlOrdOccsCount - .lOccs
'            End If
'        End With
'    Next
'    For Each objLinkToRemove In objWordToRemove.colPrevLinks
'        With objLinkToRemove
'            ' For a link to an ordinary word, update the total weight to remove both weightings
'            ' ### Needs to be twice because we're counting both backward and forward links
'            ' ### If the link is to a control word, then only 1 weight needs to be removed
'            If .lLinkNo >= conFirstWordObjNo Then
'                mlOrdOccsCount = mlOrdOccsCount - .lOccs * 2
'            Else
'                mlOrdOccsCount = mlOrdOccsCount - .lOccs
'            End If
'        End With
'    Next
    
    ' Remove word from the collection Vocab owns.
    mcolVocabWords.Remove objWordToRemove
    
    ' Update the internal statistics
    mlWordsCount = mlWordsCount - 1
    
    Set Remove = colWordsAffected
    
End Function

' Resolve: Given any object number, resolve this to a word reference (pointer)
' Resolve: Catch any errors, so it's obvious if something goes wrong
' Arguments: Object Number
' Returns: Word reference
Public Function Resolve(objno As Long) As clsWord

    If objno >= conFirstWordObjNo Then
        Set Resolve = mcolVocabWords.Find(mcolLookUp(objno))
    Else
        Set Resolve = mcolControlWords.Find(mcolLookUp(objno))
    End If
    
End Function

' Save: Save the member data for the words object to a file
' Arguments: fileno - open fileno to save to
' Returns: non-zero if error
Public Function Save(ByRef fileno As Integer) As Long

    On Error GoTo ErrorHandler
    
    Dim objWordToSave As clsWord
    
    ' Save member variables
    Put fileno, , mlWordsCount
    Put fileno, , mlOrdOccsCount
    Put fileno, , mlConOccsCount
    Put fileno, , mlNextObjNo
    
    ' File checking
    Put fileno, , MAGICBYTE

    ' Save control words
    SaveCtrlWords fileno
    
    ' Save vocabulary words
    For Each objWordToSave In mcolVocabWords.mcolWords
        ' Save this Word
        Save = objWordToSave.Save(fileno): If Save Then gobjLog.Log "Problem saving word: " + objWordToSave.strWord, Gen: Exit Function
    Next
    
    ' Signal everything went okay
    Save = 0
    
    Exit Function
    
ErrorHandler:
    Save = Err.Number
    
End Function

' Remove all references contained in all word objects (vocab and control) in order to allow objects to terminate
' Trust: N/A
' Arguments: NONE
' Returns: NONE
Public Sub TearDown()

    ' Remove references in this object
    Set mobjOwnerActor = Nothing

    Dim objWord As clsWord
    
    ' Tear down vocab words
    For Each objWord In mcolVocabWords.mcolWords
        objWord.TearDown
    Next
    
    ' Tear down control words
    For Each objWord In mcolControlWords.mcolWords
        objWord.TearDown
    Next

End Sub

' --------------------------------------------------------------------------------
' ********************************************************************************
' --------------------------------------------------------------------------------

' LoadCtrlWords: Load the details of the control words from a file
' Trust: HIGH
' Arguments: lLoadFileNo
' Returns: non-zero if error
Private Function LoadCtrlWords(nLoadFileNo As Integer) As Long

    On Error GoTo ErrorHandler

    ' Load the control words in order, from the lowest object number to the highest
    Dim WordToLoadObjNo As Long, objWordToLoad As clsWord
    For WordToLoadObjNo = conSOSObjNo To conEOSObjNo
    
        ' Find the word object associated with the object number
        Set objWordToLoad = Resolve(WordToLoadObjNo)
        
        ' Load the control word details
        objWordToLoad.Load nLoadFileNo, Me
        
        ' Update the vocabulary display list to reflect details
        frmVocab.Update objWordToLoad
        
    Next
    
    ' Signal no error
    LoadCtrlWords = 0
    
ErrorHandler:
    
    LoadCtrlWords = Err.Number
    
End Function

' NextObjNo: Return the next available object number and upgrade it
' Args: None
' Returns: Next available object number
Private Function NextObjNo() As Long
    NextObjNo = mlNextObjNo
    mlNextObjNo = mlNextObjNo + 1
End Function

' SaveCtrlWords: Save the control words to file
' Trust: LOW
' Arguments: lSaveFileNo
' Returns: NONE
Private Sub SaveCtrlWords(nSaveFileNo As Integer)

    ' Save the control words in order, from the lowest object number to the highest
    Dim WordToSaveObjNo As Long, objWordToSave As clsWord
    For WordToSaveObjNo = conSOSObjNo To conEOSObjNo
        ' Find the word object associated with the object number
        Set objWordToSave = Resolve(WordToSaveObjNo)
        ' Save the control word details
        objWordToSave.Save nSaveFileNo
    Next
    
End Sub

' Read-write properties
Public Property Get lConOccsCount() As Long: lConOccsCount = mlConOccsCount: End Property
Public Property Let lConOccsCount(lConOccsCount As Long): mlConOccsCount = lConOccsCount: End Property
Public Property Get lOrdOccsCount() As Long: lOrdOccsCount = mlOrdOccsCount: End Property
Public Property Let lOrdOccsCount(lOrdOccsCount As Long): mlOrdOccsCount = lOrdOccsCount: End Property
Public Property Get lWordsCount() As Long: lWordsCount = mlWordsCount: End Property
Public Property Let lWordsCount(lWordsCount As Long): mlWordsCount = lWordsCount: End Property

' Read only properties
Public Property Get colVocabWords() As clsWords: Set colVocabWords = mcolVocabWords: End Property

Private Sub Class_Initialize()

    mlConOccsCount = 0
    mlOrdOccsCount = 0
    mlNextObjNo = conFirstWordObjNo     ' Word object numbers will start at #0 on first call of NextObjNo()
    Set mcolVocabWords = New clsWords
    
    Set mcolLookUp = New clsLookUp
    
    ' Setup control 'words'
    Set mcolControlWords = New clsWords
    
    ' Setup EOS 'word' with appropriate object number and other details
    Dim objEOSWord As clsWord: Set objEOSWord = New clsWord
    objEOSWord.Create conEOSString, conEOSObjNo, Me
    ' Add EOS to the collection of control words for future use
    mcolControlWords.Add objEOSWord
    ' Add EOS to lookup
    mcolLookUp.Add conEOSObjNo, conEOSString
    ' Display the EOS control word in the Vocabulary listview
    frmVocab.AddWord objEOSWord
    
    ' Setup SOS 'word' with appropriate object number and other details
    Dim objSOSWord As clsWord: Set objSOSWord = New clsWord
    objSOSWord.Create conSOSString, conSOSObjNo, Me
    ' Add SOS to the collection of control words for future use
    mcolControlWords.Add objSOSWord
    ' Add SOS to lookup
    mcolLookUp.Add conSOSObjNo, conSOSString
    ' Display the SOS control word in the Vocabulary listview
    frmVocab.AddWord objSOSWord
    
End Sub

Private Sub Class_Terminate()
    Debug.Print "clsVocab terminated"
End Sub
