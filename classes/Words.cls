VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWords"
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

' clsWords
' This class is a collection class for words
' It stores only non-control words (not, for instance, EOS or SOS control words)

Option Explicit

Public mcolWords As clsCollection

' Add: Only allow Words to be added to the collection, and add them using the Word number as the key
' Add: I can't use the word itself as the Word because keys aren't case sensitive
' Arguments: pnode - word to add.  bUseKey - optional argument to use a key (node number) or not
' Arguments: bBefore - optional argument to place word at the front of the wrapped collection
' Returns: NONE
Public Function Add(pnode As clsWord, Optional bUseKey = True, Optional bBefore As Boolean = False)

    Dim key As String
    Dim nReturn As Integer
    
    ' Define key as string of Word number
    ' # If bUseKey Then key = CStr(pnode.nNoNo)
    If bUseKey Then key = pnode.strWord

    ' XXX Disabled alphabetical adding XXX
    ' XXX I suspect this is a hangover of the days before VocabListView XXX
    ' XXX Needs to be disabled for successful passing of Niall's response from MakeSentence
    ' XXX If need to re-enable, do something like add sort property to this class XXX
    ' If this is the first Word added, we don't need the alphabetical ordering
    ' If mcolWords Is Nothing Then mcolWords.Add pnode, key: Exit Function
'    ' Add new Word in correct alphabetical order
'    For Each inode In mcolWords
'        ' If we come to a word which should be after the new word, then add the Word before this and pop
'        If StrComp(inode.strWord, pnode.strWord) > 0 Then mcolWords.Add pnode, key, CStr(inode.nNoNo): Exit Function
'    Next

    ' Otherwise, if we haven't found a later word, add the new Word
    ' Add on the front or on the end, as indicated by the Before flag
    If bUseKey Then
        If bBefore Then nReturn = mcolWords.Add(pnode, key, 1) Else nReturn = mcolWords.Add(pnode, key)
        ' If bBefore Then mcolWords.Add pnode, key, 1 Else mcolWords.Add pnode, key
        
        ' Error checking
        If nReturn Then MsgBox "clsWords:Add called clsCollection:Add with error " + CStr(nReturn) + " adding object " + CStr(pnode.nNoNo) + " using key " + key, vbCritical
        
    Else
    
        If bBefore Then mcolWords.Add pnode, , 1 Else mcolWords.Add pnode
        
    End If
    
End Function

' Count: Return the number of words in this words collection
' Arguments: NONE
' Returns: long
' Trust: NONE
Public Function Count() As Long

    Count = mcolWords.Count
    
End Function

' Find: Find on the search word and return the appropriate Word
' Arguments: sword - search word
' Returns: found word or Nothing if not found
Public Function Find(sword As String) As clsWord

    Set Find = mcolWords(sword)
    
End Function

' GetConRnd: Return a random word using the weighting of contextual nodes in the collection
' Trust: HIGH
' Arguments: NONE
' Returns: word - randomly selected contextual word
Public Function GetConRnd() As clsWord

    Dim strMsg$, strLog$
    
    ' Count the contextual links contained in the collection at this moment, for later use
    Dim lTotalConLinks As Long
    lTotalConLinks = CountConLinks()
    
    ' Sanity check - If there's no vocab yet then exit the function as appropriate
    If lTotalConLinks = 0 Then Set GetConRnd = Nothing: Exit Function
    
    ' Elsewise, get a random Word and return the word from that
    ' First choose the random occurence number to pick a word out of the weighted collection
    Dim lRndValue As Long, objNode As clsWord
    lRndValue = rndint(1, lTotalConLinks)
    
    ' Now actually find out which word in the collection this number represents
    For Each objNode In mcolWords
        With objNode
            lRndValue = lRndValue - .colConLinks.lOccs
            If lRndValue <= 0 Then Set GetConRnd = objNode: Exit Function
        End With
    Next
   
    ' Will never reach this point in the program, since lRndValue will never be greater than no of contextual links
    ' *** So let the user know to contact me!  But return a value to continue - since this is non-fatal
    strLog = "clsWords:GetConRnd generated as error with lTotalConLinks=" + CStr(lTotalConLinks) + " lRndValue=" + CStr(lRndValue)
    strMsg = "An error has occured in clsWords:GetConRnd.  Please e-mail " + EMAIL + " with " + APPNAME + "'s logfile."
    gobjLog.Log strLog, Warn
    MsgBox strMsg, vbCritical
    Set GetConRnd = Nothing
    
End Function

' GetOrdRnd: Return a random word from the collection, weighting for instances.
' Arguments: lOccs - argument to signify occurences, which is upper limit of random bounds to choose word
' Arguments: lOccs - needs to be an arg since it may be for all vocab, or for a temporary collection (eg contextual)
' Returns: reference to word chosen
Public Function GetOrdRnd(lOccs As Long) As clsWord

    ' Sanity check - If there's no vocab yet then exit the function as appropriate
    If lOccs = 0 Then Set GetOrdRnd = Nothing: Exit Function
    
    ' Elsewise, get a random Word and return the word from that
    ' First choose the random occurence number to pick a word out of the weighted collection
    Dim lRndValue As Long, objNode As clsWord
    lRndValue = rndint(1, lOccs)
    
    ' Now actually find out which word in the collection this number represents
    For Each objNode In mcolWords
        With objNode
            ' We only need to look at both link lists to get the number of links
            ' Next and prev will not necessarily be the same, if deletion has taken place, so delete both.
            lRndValue = lRndValue - .colNextLinks.lOccs
            lRndValue = lRndValue - .colPrevLinks.lOccs
            If lRndValue <= 0 Then Set GetOrdRnd = objNode: Exit Function
        End With
    Next
   
    ' Should never reach this point in the program, since val should never be greater than upper bound of objects
    ' It's up to the calling function to distinguish between this error and a 'zero occurences' one, above.
    Set GetOrdRnd = Nothing
    
End Function

' NewEnum must return the IUnknown interface of a
' collection's enumerator.
' (This copied from Visual Basic Books Online to allow 'for each ... in' to be used)
' Bit of a hack really - needs Tools - Procedure Attributes to have ID=-4 and this procedure hidden
'Public Function NewEnum() As IUnknown
'    Set NewEnum = mcolWords.[_NewEnum]
'End Function

' Remove Word from collection by Word number key
Public Function Remove(n As clsWord)

    mcolWords.Remove n.strWord
    
End Function

' --------------------------------------------------------------------------------
' ********************************************************************************
' --------------------------------------------------------------------------------
' CountConLinks: Counts the number of contextual links contained in the collection
' Trust: HIGH
' Arguments: NONE
' Returns: long - count of collection contextual links
Private Function CountConLinks() As Long

    Dim objIWord As clsWord
    CountConLinks = 0
    
    ' Iterate through each word and add it's contextual link count
    For Each objIWord In mcolWords
        With objIWord
            CountConLinks = CountConLinks + .colConLinks.lOccs
        End With
    Next
End Function

Private Sub Class_Initialize()

    ' Apparantly, creating the object when the class its in is initialized results in less overhed
    ' than if the object were declared as new
    Set mcolWords = New clsCollection
    mcolWords.Initialise vbObject
    
End Sub

Private Sub Class_Terminate()
    ' Debug.Print "DESTROYED words class"
End Sub
