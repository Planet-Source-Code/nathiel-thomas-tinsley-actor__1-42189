VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsActor"
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

' This class will eventually contain all the details for
' implementing a collection for Niall's vocabulary using Words
Option Explicit

' Constants
Private Const OFFSET& = vbObjectError + &H200
Private Const errLoadFileHigher& = OFFSET + &H100   ' Load file version is higher than this program!

' Members
Private mnMajor, mnMinor, mnRevision As Integer     ' Version number of the Actor program that created this Actor

Public mobjLog As clsLog                            ' This actor's conversation log.  Referencing to log's vars done through actor object.

Private mbIsInit As Boolean                         ' Signals whether the actor has been initialised

Private mstrName As String                          ' The name of this actor
Private mbChangedFlag As Boolean                    ' Flag for whether any aspect of the actor (including vocab) has changed from last save
Private mbBackwardLinking As Boolean                ' Flag if sentences should be constructed by linking backward from the seed word as well as forward
                                                    ' ***NOT CURRENTLY AVAILABLE TO BE SET AS AN OPTION***
Private mbContextual As Boolean                     ' Flag to set whether the initial reply word chosen by Niall attempts to use contextual information
Private mnMinConSize As Integer                     ' Lowest size for a word to be considered a contextual link.

Private mbExclude As Boolean                        ' Contextual exclusion

Private mbLimReplyLength As Boolean                 ' Curtailing replies?
Private mnPrefMaxReplyLength As Integer             ' Once a reply sentence reaches this length we want to end it ASAP
Private mfTermProportionReq As Single               ' When trying to end an over-long sentence, this is how big the proportion of EOS or SOS
                                                    ' needs to be for the word to be regarded as a valid EOS/SOS.
Private mcolActorConOldWords As clsWords            ' Contextual words in Actor's second most recent reply
Private mcolActorConNewWords As clsWords            ' Contextual words in Actor's most recent reply
Private mcolUserConWords As clsWords                ' Contextual words in user's reply

Public mcolExclusions As clsExclusions              ' Pointer to the contextual exclusions relevant to this actor

Private trnodes As clsVocab                         ' clsVocab which this Actor uses for utterances

' Analyze: Analyze an array of string tokens that comes in.
' Analyze: At the moment, these means adding and increment words as appropriate, and creating contextual links
' Arguments: arrTokens - array of strings which represents the words in the user's input sentence
' Returns: NONE
Public Sub Analyze(ByRef arrTokens As Variant)

    Debug.Assert mbIsInit
    
    ' Ingesting the user's input
    Dim colWords As clsWords
    Set colWords = AddTokens(arrTokens)

    ' Contextual link both ways, from user->actor and actor->user
    AddConLinks mcolActorConOldWords, mcolUserConWords
    AddConLinks mcolUserConWords, mcolActorConOldWords
    
    ' Capture the contextual words used by the user
    Set mcolUserConWords = GetCntxtlWords(colWords)
    
End Sub

' Import: Take a textfile and add it as vocabulary with ordinary and contextual links
' Trust:
' Arguments: strFileName
' Returns: boolean - true if an error has occured
Public Function Import(nFileNo As Integer) As Boolean

    Debug.Assert mbIsInit
    
    On Error GoTo ErrorHandler

    ' Analyse input file line by line
    Dim i As Integer: i = 0   ' ***DEBUG LINE
    Dim strImportLine As String
    Dim colWords As clsWords, colOldImportConWords As clsWords, colNewImportConWords As clsWords

    While Not EOF(nFileNo)

        strImportLine = basGeneral.InputLine(nFileNo)
        
        ' strip out invalid characters
        strImportLine = basGeneral.StripNonPrint(strImportLine)
        
        ' Add the words in this sentence to the vocabulary
        While strImportLine <> ""
            Set colWords = AddTokens(AnalyzeSentence(strImportLine))
            
            ' Store the contextual words from the previous sentence and get the new ones from this sentence
            Set colOldImportConWords = colNewImportConWords
            Set colNewImportConWords = GetCntxtlWords(colWords)
            
            ' Contextual link the analyzed sentence and the previous sentence
            AddConLinks colOldImportConWords, colNewImportConWords
            AddConLinks colNewImportConWords, colOldImportConWords
            
            ' Debug, keep displaying how many lines have been imported, since it takes a long time
            i = i + 1
            If i Mod 20 = 0 Then Debug.Print "<"; Format(i); " lines imported>"
        Wend
    Wend
    
    frmMonitor.Output "<" + Format(i) + " lines imported>", Low
    Exit Function

ErrorHandler:

    MsgBox "Error importing text", vbExclamation

End Function

' Initialise: Set up the actor - this has to be called before anything else
' Trust:
' Arguments:
' Returns: NONE
Public Sub Initialise(ByRef colExclusions As clsExclusions)

    Set mcolExclusions = colExclusions
    
    mbIsInit = True
    
End Sub

' Load: Load the Actor
' Trust: All the information loaded from the saved file is within acceptable bounds
' Arguments: filename - name of file to load from
' Returns: non-zero if error
Public Function Load(ByRef nFileNo As Integer) As Long
    
    On Error GoTo ErrorHandler
    
    Debug.Assert mbIsInit
    
    Dim colWords As clsWords, objWord As clsWord
    Dim lLen&
    Dim yCheck As Byte
    Dim lReturnCode&
    Dim nMyVersion%
    
    gobjLog.Log "<start load>", Info
    
    ' Load Actor variables
    Get nFileNo, , mnMajor
    Get nFileNo, , mnMinor
    Get nFileNo, , mnRevision
       
    ' Construct ordinal version numbers
    gnVersion = GetVersion(mnMajor, mnMinor, mnRevision)
    nMyVersion = GetVersion(App.Major, App.Minor, App.Revision)
    
    ' Sanity - check this program can handle the version of the file we are loading
    If nMyVersion < gnVersion Then Err.Raise errLoadFileHigher

    If gnVersion >= 14 Then
        Get nFileNo, , lLen
        mstrName = Input(lLen, nFileNo)
    Else
        Get nFileNo, , mstrName
    End If
    
    Get nFileNo, , mbContextual
    Get nFileNo, , mnMinConSize
    
    ' From 0.2.0 onwards contextual exclusion was introduced
    If gnVersion >= 20 Then Get nFileNo, , mbExclude
            
    ' From 0.1.7 onwards backward linking information was moved from an option (global) to a property (actor specific)
    ' which is what it should have been in the first place :-P
    If gnVersion >= 17 Then
        Get nFileNo, , mbBackwardLinking
    End If
    
    ' From 0.1.7 onwards, a feature to curtail long sentences was introduced
    ' Need to save status of this feature, preferred reply length and threshold at which a word is considered a SOS or EOS terminator
    If gnVersion >= 17 Then
        Get nFileNo, , mbLimReplyLength
        Get nFileNo, , mnPrefMaxReplyLength
        Get nFileNo, , mfTermProportionReq
    End If
    
    ' From 0.1.4 onwards logging was introduced, so only from this version load logging path/file title information
    If gnVersion >= 14 Then
        Get nFileNo, , lLen
        mobjLog.mstrLogPath = Input(lLen, nFileNo)
        Get nFileNo, , lLen
        mobjLog.mstrLogTitle = Input(lLen, nFileNo)
    End If
    
    ' Check integrity of file
    If gnVersion >= 20 Then
        Get nFileNo, , yCheck
        
        If yCheck <> MAGICBYTE Then
            gobjLog.Log "Load file loses integrity loading clsActor variables", Error
            Err.Raise errMagicMisplaced
        End If
        
    End If
    
    ' Load all Words
    lReturnCode = trnodes.Load(nFileNo): If lReturnCode Then Err.Raise lReturnCode
          
    ' If set, purge any excluded context words from the incoming actor
    If mbExclude Then
    
        Set colWords = trnodes.PurgeConExcluded()
        
        For Each objWord In colWords.mcolWords
            frmVocab.Update objWord
        Next
        
    End If
    
    ' Update version numbers of the newly loaded file so they are most current when the actor is resaved
    mnMajor = App.Major: mnMinor = App.Minor: mnRevision = App.Revision
    
    ' Indicate no aspect of the actor has yet been changed
    mbChangedFlag = False
    
    gobjLog.Log Format(trnodes.lWordsCount) + " words and " + CStr(trnodes.lOrdOccsCount) + " ordinary occurences.", Info
    gobjLog.Log "<end load>", Info
    
    Exit Function
    
ErrorHandler:

    Load = Err.Number

' ### If we tried to load a non-Actor file this error is _usually_ thrown. ###
' ### However, this is not a reliable indicator ###
' ### A better way to handle such a situation would be to embed a magic string at the start of an actor file for identification purposes ###
' ### But this has legacy issues ###
' Select Case Err.Number
'    Dim nul As Integer
' (Variable uses a type not supported in Visual Basic)
' Usually occurs when the file selected isn't a valid Niall vocab file
'    Case 458
'        nul = MsgBox("Not an " + conAppName + " vocabulary file!", vbExclamation + vbOKOnly, "Invalid file")
'        Exit Function
'    End Select
    
End Function

' Reply: Construct an output sentence using this actor's vocab, contextual links, etc.
' Arguments: NONE
' Returns: String - well formed reply
Public Function Reply() As String

    Debug.Assert mbIsInit

    ' Get a response by getting the response collection, storing the potential contextual words, and building the string
    ' Words collection we've used to store Niall's sentence gets destroyed when colSentenceWords goes out of scope
    ' !!! However, if a collection is created in the function and passed back, it never seems to get destroyed !!!
    ' !!! Perhaps it never goes out of scope (effectively global) !!!
    Dim objSeedWord As clsWord  ' This word (once selected) acts as the word which builds the sentence through linking
    Dim colSentenceWords As clsWords
    
    ' First, see if there is a viable contextual seed word if contextual linking has been selected
    Set objSeedWord = Nothing
    If mbContextual Then Set objSeedWord = GetConSeedWord()
    
    ' If the contextual operations haven't generated a seed word, then select a random word from the entire vocabulary
    If objSeedWord Is Nothing Then
        Set objSeedWord = trnodes.GetOrdRnd
        
        ' Update monitor information
        If Not objSeedWord Is Nothing Then frmMonitor.Output "Replying using random seed word: " + objSeedWord.strWord, Medium

    End If
    
    ' Generate the reply string using the selected words
    Set colSentenceWords = MakeSentenceWords(objSeedWord)
    
    ' Transfer the actors last contextual words to storage and capture the contextual words of this reply
    Set mcolActorConOldWords = mcolActorConNewWords
    Set mcolActorConNewWords = GetCntxtlWords(colSentenceWords)

    ' Return reply string
    Reply = MakeSentenceString(colSentenceWords)
    
End Function

' Reweight: Reweight a link for a specified word
' Trust:
' Arguments: objWordForLink, objLinkToReweight, lNewWeight
' Returns: clsLink - link reweighted
Public Function Reweight(objWordForLink As clsWord, objLinkToReweight As clsLink, lNewWeight As Long) As clsLink

    Debug.Assert mbIsInit
    
    ' Set the appropriate contextual links entry to the new weighting
    objWordForLink.colConLinks.Reweight objLinkToReweight, lNewWeight
    
    ' Signal the actor has definitely changed since the last save
    mbChangedFlag = True

End Function

' Save: Save the Actor
' Arguments: fileno
' Returns: non-zero if error
Public Function Save(ByRef fileno As Integer) As Long
    
    On Error GoTo ErrorHandler
    
    Debug.Assert mbIsInit
    
    ' Start Actor stuff
    Put fileno, , mnMajor
    Put fileno, , mnMinor
    Put fileno, , mnRevision
    Put fileno, , Len(mstrName)
    Put fileno, , mstrName
    Put fileno, , mbContextual
    Put fileno, , mnMinConSize
    Put fileno, , mbExclude
    Put fileno, , mbBackwardLinking
    Put fileno, , mbLimReplyLength
    Put fileno, , mnPrefMaxReplyLength
    Put fileno, , mfTermProportionReq
    
    With mobjLog
        Put fileno, , Len(.mstrLogPath)
        Put fileno, , .mstrLogPath
        Put fileno, , Len(.mstrLogTitle)
        Put fileno, , .mstrLogTitle
    End With
    ' End Actor stuff
    
    ' File checking
    Put fileno, , MAGICBYTE
    
    Save = trnodes.Save(fileno): If Save Then gobjLog.Log "Vocabulary failed to save in Actor", Error: Exit Function
    
    ' Indicate no aspect of this actor has now changed since the last save
    mbChangedFlag = False
    
    ' Indicate no save errors
    Save = 0
    
    Exit Function
    
ErrorHandler:
    Save = Err.Number
    
End Function

' TearDown: Reset all the references in the Actor to nothing to allow it to terminate
Public Sub TearDown()

    ' Tear down each word in the vocabulary
    trnodes.TearDown
    
End Sub

' ********************************************************************************
' --------------------------------------------------------------------------------
' ********************************************************************************

' AddConLinks: Link each word in a source collection to each word in a target collection for this actor
' AddConLinks: At the end, removes all words held in mcolActorConWords and zaps colTarget (passed collection argument)
' Arguments: colSource - Collection of source words, colTarget - Collection of target words
' Returns: NONE
Private Sub AddConLinks(ByRef colSource As clsWords, ByRef colTarget As clsWords)

    Dim objConSourceWord As clsWord, objConTargetWord As clsWord
    Dim objDebugLink As clsLink
    
    ' Sanity check - see if there's anything has been stored from Niall's output to actually link with
    ' If not, exit the sub
    If colSource Is Nothing Then
        frmMonitor.Output "AddConLinks: Nothing to link with!", High
        Exit Sub
    End If
    If colTarget Is Nothing Then
        frmMonitor.Output "AddConLinks: Nothing to link to!", High
        Exit Sub
    End If
    
    ' With each contextual source word, add each contextual target word as a link
    For Each objConSourceWord In colSource.mcolWords
        For Each objConTargetWord In colTarget.mcolWords
            With objConSourceWord
            
                Set objDebugLink = .AddConLink(objConTargetWord.nNoNo, 1)
                frmMonitor.Output "AddConLinks: " + objConSourceWord.strWord + " => " + trnodes.Resolve(objDebugLink.lLinkNo).strWord, High

            End With
        Next
    Next
    
End Sub

' Addtokens: Add an array of string tokens to the actor's vocabulary
' Arguments: word - array of tokens to add.  Array should have postfix of "" to be valid
' Returns: Words - a collection of the words that have just been added
Private Function AddTokens(ByRef word As Variant) As clsWords

    ' Debug.Print "<start adding tokens>"
    
    Dim i As Integer: i = 0
    Set AddTokens = New clsWords
    
    ' Processing the token list involves two passes
    ' Pass 1 : Add tokens as nodes, until we reach the empty string marker
    ' Mark the start and end of the sentence in the node array
    ' This is done using SOS and EOS 'words'
    ReDim node(1) As clsWord: Set node(0) = trnodes.Resolve(conEOSObjNo)
    While word(i) <> ""
        ReDim Preserve node(i + 2)
        Set node(i + 1) = trnodes.Add(CStr(word(i)))
        i = i + 1
    Wend
    
    ' This is a bodge way to signal something has changed when the sentence is parsed
    ' A better way would be to more cleanly detect when nothing has been said, when the change flag shouldn't be set
    If i <> 0 Then mbChangedFlag = True
    
    ' Mark the EOS
    ReDim Preserve node(i + 2): Set node(i + 1) = trnodes.Resolve(conSOSObjNo)
    
    ' Pass 2 : Add links to new Words.
    ' DON'T FORGET FOR HISTORICAL REASONS, THE NODE LIST TRAVELS FROM THE END OF THE SENTENCE TO THE START
    ' First, link the first node (the EOS node) to the last word
    i = 0
    node(i).AddCtrlPrevLink node(i + 1).nNoNo
    
    ' Second, add the word nodes
    i = i + 1
    While node(i).nNoNo <> conSOSObjNo
        ' Don't forget back to front, so the previous link node is the next in the node list and vice-versa
        node(i).AddOrdPrevLink node(i + 1).nNoNo
        node(i).AddOrdNextLink node(i - 1).nNoNo
         ' Add each node to the return collection.  Doesn't matter it's back to front.
        AddTokens.Add node(i), False
        i = i + 1
    Wend
    
    ' Third, link the start of sentence node
    node(i).AddCtrlNextLink node(i - 1).nNoNo
    
    ' Destroy the word array now we've now finished with it
    Set word = Nothing
    
    ' Debug.Print "<end adding tokens>"
    
End Function

' GetConSeedWord: Attempt to return a contextual seed word from the user's candidate contextual words
' Arguments: NONE
' Returns: word - contexual seed word chosen, Nothing if none has been chosen
Private Function GetConSeedWord() As clsWord

    Dim objSubjectWord As clsWord  ' The word picked out from the user's input to pick a contextual link from

    ' Sanity check - there will be no suitable word if the user's sentence contained no contextual words
    If mcolUserConWords Is Nothing Then Set GetConSeedWord = Nothing: Exit Function
    
    ' Choose a random contextual word from the collection
    Set objSubjectWord = mcolUserConWords.GetConRnd()
    
    ' Sanity check - Word:GetConRnd() can return Nothing if none of the words in the collection have any contextual links
    If objSubjectWord Is Nothing Then Set GetConSeedWord = Nothing: Exit Function
    
    ' Update monitor information
    frmMonitor.Output objSubjectWord.strWord & " chosen as the subject word from the last sentence.", Medium
    
    ' Get a random contextual link
    Dim objSeedLink As clsLink
    Set objSeedLink = objSubjectWord.colConLinks.GetRandom
    
    ' Assertion - since we know there's at least one word with a contextual link from GetConRnd,
    ' objSeedLink should always return a valid link
    Debug.Assert Not objSeedLink Is Nothing
    
    ' Return the seed random contextual word
    Set GetConSeedWord = trnodes.Resolve(objSeedLink.lLinkNo)

    ' Update monitor information
    frmMonitor.Output "Replying using " & GetConSeedWord.strWord & " as a word associated with " & objSubjectWord.strWord, Medium
    
End Function

' GetCntxtlWords: Extract all the words which are contextual candidates from the input collection
' GetCntxtlWords: At the moment, the filter is word length, as defined in the Options object
' Arguments: colSentenceWords - collection of words to grab the contextual words from
' Returns: Words - collection of contextual candidate words
Private Function GetCntxtlWords(colSentenceWords As clsWords) As clsWords
    
    Set GetCntxtlWords = New clsWords
    Dim strConForm As String
    
    ' Don't parse if there's nothing to parse
    If colSentenceWords Is Nothing Then Set GetCntxtlWords = Nothing: Exit Function
    
    ' For each word in the collection, add it to the contextual collection if it meets the criteria
    Dim objNode As clsWord
    For Each objNode In colSentenceWords.mcolWords
    
        ' Preparation - reduce word to its contextual form (minus nonalphanumeric leading and trailing chars)
        strConForm = objNode.ConWord()
    
        ' First pass - word needs to be long enough
        If Len(strConForm) >= mnMinConSize Then
        
            ' Second pass - if we are using the exclusion list, word cannot be on that list.
            If mbExclude Then
            
                If mcolExclusions(strConForm) = "" Then GetCntxtlWords.Add objNode, False: Debug.Print "Added: "; objNode.strWord
                
            ' If we're not using exclusing, then just add the word to the contextual list
            Else
            
                GetCntxtlWords.Add objNode, False: Debug.Print "Added: "; objNode.strWord
            
            End If
            
        End If
        
    Next
    
End Function

' MakeSentenceString: Take a collection of words and turn it into a properly formatted sentence
' MakeSentenceString: !!! Destroys the collection of words passed to it !!!
' Arguments: colSentenceWords - collection to turn into the sentence
' Returns: string - the sentence
Private Function MakeSentenceString(colSentenceWords As clsWords) As String

    If colSentenceWords Is Nothing Then MakeSentenceString = "Please say something!": Exit Function
    ' Add the actual words
    Dim node As clsWord
    For Each node In colSentenceWords.mcolWords
        MakeSentenceString = MakeSentenceString + node.strWord + " "
    Next
    
    ' Snip the last space off, append full stop and capitalise
    Dim s As String, lastchar As String
    MakeSentenceString = Left$(MakeSentenceString, Len(MakeSentenceString) - 1)  ' Snip the extra space off
    lastchar = VBA.Right$(MakeSentenceString, 1)
    If lastchar <> "?" And lastchar <> "!" Then MakeSentenceString = MakeSentenceString + "."
    ' Extract first character of sentence, capitalise and stick it back on the start
    s = Left$(MakeSentenceString, 1)
    s = StrConv(s, vbUpperCase)
    Mid$(MakeSentenceString, 1, 1) = s
    
End Function

' MakeSentenceWords: Make up the sentence nodes from the available vocabulary using a given word
' Arguments: objSeedWord - initial word to use in forward/backward linking to create sentence
' Returns: A collection of sentence nodes in the right order
Private Function MakeSentenceWords(objSeedWord As clsWord) As clsWords

    ' Debug.Print "BEGIN MakeSentenceWords()"
    Dim objCurrWord As clsWord, objCtrlLink As clsLink
    Set MakeSentenceWords = New clsWords
    
    ' Sanity check - if we've been given nothing as a word then no vocab exists and we can't make a sentence
    If objSeedWord Is Nothing Then Set MakeSentenceWords = Nothing: Exit Function
    
    ' Add the initial word into the return collection
    MakeSentenceWords.Add objSeedWord, False

    ' Set up subsequent words until we have a complete reply
    ' ~~~ We also want to bear in mind that once the length of the reply exceeds the preferred length, any word which
    ' ~~~ has a significant number of links which start/end a sentence will become the start/end word, to stop the
    ' ~~~ sentence going on too long
    
    ' First, get forward links
    Set objCurrWord = objSeedWord.GetRandomLink(forward)
    Do While Not objCurrWord Is Nothing
        
        ' End the sentence if we've reached an EOS node
        If objCurrWord.nNoNo = conEOSObjNo Then Exit Do
        
        ' Add the current word on the end of the sentence being constructed
        MakeSentenceWords.Add objCurrWord, False
        
        ' See if the current word can end the sentence if we've reached the preferred reply length if this is enabled
        If mbLimReplyLength Then
            If MakeSentenceWords.Count >= mnPrefMaxReplyLength Then
                If objCurrWord.IsValidTerminator(forward, mfTermProportionReq) Then
    
                    frmMonitor.Output "Reply sentence capped linking forward", High
                    Exit Do
    
                End If
            End If
        End If
                
        ' Otherwise, randomly select the next word in the sequence
        Set objCurrWord = objCurrWord.GetRandomLink(forward)
        ' If Not objCurrWord Is Nothing Then Debug.Print objCurrWord.strWord
        
    Loop
    
    ' Second if we've got this option selected, get backward links
    If mbBackwardLinking Then
    
        ' Check if the seed word can act as a terminator if we've reached the preferred reply length
        Set objCurrWord = objSeedWord
        If mbLimReplyLength Then
            If MakeSentenceWords.Count >= mnPrefMaxReplyLength Then
            
                If objCurrWord.IsValidTerminator(backward, mfTermProportionReq) Then
                
                    frmMonitor.Output "Reply sentence capped linking backward", High
                    Exit Function
                    
                End If
                
            End If
        End If
    
        ' Now build the words in the sentence that will preceed the seed word
        Set objCurrWord = objSeedWord.GetRandomLink(backward)
    
        Do While Not objCurrWord Is Nothing
    
            ' End sentence building if we've reached an EOS node
            If objCurrWord.nNoNo = conSOSObjNo Then Exit Do
        
            ' Add the current word at the start of the sentence being constructed
            MakeSentenceWords.Add objCurrWord, False, True
            
            ' See if the current word can start the sentence if we've reached the preferred reply length
            If mbLimReplyLength Then
                If MakeSentenceWords.Count >= mnPrefMaxReplyLength Then
                    If objCurrWord.IsValidTerminator(backward, mfTermProportionReq) Then
        
                        frmMonitor.Output "Reply sentence capped linking backward", High
                        Exit Do
        
                    End If
                End If
            End If
    
           ' Otherwise, randomly select the preceeding word in the sequence
            Set objCurrWord = objCurrWord.GetRandomLink(backward)
            
            ' If Not objCurrWord Is Nothing Then Debug.Print objCurrWord.strWord
        Loop

    End If
    
    ' Debug.Print "END MakeSentenceWords()"
    
End Function

' Print out the contents of the vocab collection class using the hidden for..each enumerator
'Public Sub DebugPrintActor()
'    Dim pnode as clsWord
'
'    Debug.Print "<start Actor>"
'
'    For Each pnode In trnodes
'        Debug.Print pnode.strWord; " - Occs:"; Format(pnode.colLinks.lOccs)
'    Next
'
'    Debug.Print "<end Actor>"
'End Sub

' Read-only properties
Public Property Get colVocab() As clsVocab: Set colVocab = trnodes: End Property

' Read-write properties
Public Property Get bBackwardLinking() As Boolean: bBackwardLinking = mbBackwardLinking: End Property
Public Property Let bBackwardLinking(bBackwardLinking As Boolean)
    ' If this property is being altered from it's existing value, then signal this.
    If mbBackwardLinking <> bBackwardLinking Then mbChangedFlag = True: mbBackwardLinking = bBackwardLinking
End Property

Public Property Get bChangedFlag() As Boolean: bChangedFlag = mbChangedFlag: End Property
Public Property Let bChangedFlag(bChangedFlag As Boolean): mbChangedFlag = bChangedFlag: End Property

Public Property Get bContextual() As Boolean: bContextual = mbContextual: End Property
Public Property Let bContextual(bContextual As Boolean)
    ' Signal if changed
    If mbContextual <> bContextual Then mbChangedFlag = True: mbContextual = bContextual
End Property

Public Property Get bExclude() As Boolean: bExclude = mbExclude: End Property
Public Property Let bExclude(bExclude As Boolean)
    ' Signal if changed
    If mbExclude <> bExclude Then mbChangedFlag = True: mbExclude = bExclude
End Property

Public Property Get nMinConSize() As Integer: nMinConSize = mnMinConSize: End Property
Public Property Let nMinConSize(nMinConSize As Integer)
    ' Sanity - contextual word size should be positive
    If nMinConSize < 1 Then Exit Property

    ' If this property is being altered from it's existing value, then signal this.
    If mnMinConSize <> nMinConSize Then mbChangedFlag = True: mnMinConSize = nMinConSize
End Property

Public Property Get bLimReplyLength() As Boolean: bLimReplyLength = mbLimReplyLength: End Property
Public Property Let bLimReplyLength(bLimReplyLength As Boolean)
    ' If this property is being altered from it's existing value, then signal this.
    If mbLimReplyLength <> bLimReplyLength Then mbChangedFlag = True: mbLimReplyLength = bLimReplyLength
End Property

Public Property Get nPrefMaxReplyLength() As Integer: nPrefMaxReplyLength = mnPrefMaxReplyLength: End Property
Public Property Let nPrefMaxReplyLength(nPrefMaxReplyLength As Integer)
    ' Sanity - we can't have a preferred reply length of less than 1 but it can be as large as we like
    If nPrefMaxReplyLength < 1 Then Exit Property
    
    ' If this property is being altered from it's existing value, then signal this.
    If mnPrefMaxReplyLength <> nPrefMaxReplyLength Then mbChangedFlag = True: mnPrefMaxReplyLength = nPrefMaxReplyLength
End Property

Public Property Get fTermProportionReq() As Single: fTermProportionReq = mfTermProportionReq: End Property
Public Property Let fTermProportionReq(fTermProportionReq As Single)
    ' Sanity - since this is a proportion must be between 0 and 1
    If fTermProportionReq < 0 Then fTermProportionReq = 0
    If fTermProportionReq > 1 Then fTermProportionReq = 1

    ' If this property is being altered from it's existing value, then signal this.
    If mfTermProportionReq <> fTermProportionReq Then mbChangedFlag = True: mfTermProportionReq = fTermProportionReq
End Property

Public Property Get strLogPath() As String: strLogPath = mobjLog.mstrLogPath: End Property
Public Property Let strLogPath(strLogPath As String)

    With mobjLog
        ' If this property is being altered from it's existing value, then signal this.
        If .mstrLogPath <> strLogPath Then mbChangedFlag = True: .mstrLogPath = strLogPath
    End With
    
End Property

Public Property Get strLogFileTitle() As String: strLogFileTitle = mobjLog.mstrLogTitle: End Property
Public Property Let strLogFileTitle(strLogFileTitle As String)

    With mobjLog
        ' If this property is being altered from it's existing value, then signal this.
        If .mstrLogTitle <> strLogFileTitle Then mbChangedFlag = True: .mstrLogTitle = strLogFileTitle
    End With
    
End Property

Public Property Get strName() As String: strName = mstrName: End Property
Public Property Let strName(strName As String): mstrName = strName: End Property

Private Sub Class_Initialize()

    ' Initialize owned objects
    Set trnodes = New clsVocab
    trnodes.Create Me
    
    Set mobjLog = New clsLog
    
    ' ***TEMP***
    mstrName = conDefName
    
    ' Default properties
    mbBackwardLinking = conDefBackwardLinking
    mnMinConSize = conDefaultMinConSize
    mbContextual = True
    
    mbExclude = conDefExclude
    
    mbLimReplyLength = conDefLimReplyLength
    mnPrefMaxReplyLength = conDefPrefMaxReplyLength
    mfTermProportionReq = conDefTermProportionReq
    
    ' Signal the Actor hasn't yet been changed
    mbChangedFlag = False
    
    ' Initialize version number, can be changed in loading later
    mnMajor = App.Major: mnMinor = App.Minor: mnRevision = App.Revision
    
    ' Signal actor is uninitialised until the user explicitly calls
    
End Sub

Private Sub Class_Terminate()

    Set mobjLog = Nothing
    
    Debug.Print "clsActor terminated"
    
End Sub
