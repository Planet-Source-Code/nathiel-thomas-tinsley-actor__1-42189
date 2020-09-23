VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLinks"
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

' clsLinks: This is the collection class for links
Option Explicit

Public mcolLinks As clsCollection  ' Has to be public in order to allow enumeration, but nothing should be added/removed directly
Private mlOccs As Long             ' Actual total number of link occurences (as opposed to # of links)
                                   ' This number is valid at all times

' Add: Add a new link to the collection
' Arguments: lLinkNo - write once linked node number
' Arguments: lOccs - number of occurences, which will either be initial or added to an existing link
' Returns: clsLink - reference to this object
Public Function Add(lLinkNo As Long, lOccs As Long) As clsLink

    ' Debug.Print "BEGIN clsLinks:Add"
    Dim objFoundLink As Object
    ' See if an identical link to the one proposed already exists
    ' # Set objFoundLink = Find(lLinkNo)
    Set objFoundLink = mcolLinks(GenWordKey(lLinkNo))

    ' IF an existing link hasn't been found then create a new link and add to the collection
    If objFoundLink Is Nothing Then
    
        ' Get link and initialise through :Create
        Dim objNewLink As New clsLink
        Set Add = objNewLink.Create(lLinkNo, lOccs)
        
        ' Add link to collection
        mcolLinks.Add objNewLink, GenWordKey(lLinkNo)
        
    ' ELSE increment the found link by the number of occurences
    Else
    
        ' Initialise
        objFoundLink.lOccs = objFoundLink.lOccs + lOccs
        Set Add = objFoundLink
        
    End If
    
    ' Update statistics on this object
    mlOccs = mlOccs + lOccs

    ' Debug.Print "END clsLinks:Add"
    
End Function

' Copy: Copy the links in a collection over to this collection
' Arguments: colSource - source collection for links
' Returns: Long of links copied
Public Function Copy(colSource As clsLinks) As Long

    ' Go through the source collection and add the existing link objects into this collection
    Dim objSourceLink As clsLink
    For Each objSourceLink In colSource.mcolLinks
        Add objSourceLink.lLinkNo, objSourceLink.lOccs
        
        ' Update the return statistic
        Copy = Copy + objSourceLink.lOccs
    Next
End Function

Public Function Count() As Long
    Count = mcolLinks.Count
End Function

' Find: Search this collection to see if a link to the specified word already exists
' Trust: LOW
' Arguments: strLinkKey - Link key to search for as a link
' Returns: clsLink - link if there's a match, Nothing if there's no match
Public Function Find(ByVal lLinkObjNo As Long) As clsLink

    Dim strLinkKey As String
    
    ' Generate key from object number
    strLinkKey = GenWordKey(lLinkObjNo)
    
    ' Sanity check - Check the string passed is a valid key
    Debug.Assert IsValidKey(strLinkKey)
    
    ' Try to retrieve a valid link from the collection
    Set Find = mcolLinks(strLinkKey)

End Function

' GetRandom: Return a random link from the collection
' Trust: HIGH, though does check if there are actually any occurences to choose from
' Arguments: NONE
' Returns: clsLink - random link or nothing if no links are available
Public Function GetRandom() As clsLink

    Dim lTotalWeight As Long:  lTotalWeight = CountWeight()
    
    ' Sanity check - if there are no occurences then there is nothing to choose
    If lTotalWeight <= 0 Then Set GetRandom = Nothing: Exit Function
    
    ' Deduct links from count until we get to the randomly selected link
    Dim a As Integer: a = rndint(0, lTotalWeight - 1)
    Dim objLink As clsLink
    For Each objLink In mcolLinks
        a = a - objLink.lOccs
        If a < 0 Then Set GetRandom = objLink: Exit Function
    Next

    ' Shouldn't ever reach this code, as a shouldn't be larger than the actual number of links
    Stop
    
End Function

' Load: Creates links and loads them with the correct values
' Arguments: nFileNo - File number to load from.  nLinks - number of links to load
' Returns: non-zero if error
Public Function Load(nFileNo As Integer, ByVal nLinks As Long) As Long

    On Error GoTo ErrorHandler

    ' Create the links required in the collection
    Dim lLinkNo As Long, lOccs As Long
    Dim objLoadedLink As clsLink
    While nLinks > 0  ' There will always be at least one link to load
        Get nFileNo, , lLinkNo
        Get nFileNo, , lOccs
        ' Add the link to this collection (create and input details)
        Set objLoadedLink = Add(lLinkNo, lOccs)
    '    Debug.Print "Added link: <word> => "; objLoadedLink.lLinkNo; " with "; objLoadedLink.lOccs; " occs."
        nLinks = nLinks - 1
    Wend
        
    ' If we got here everything is okay and we don't need to report an error
    Load = 0
    
    Exit Function
    
ErrorHandler:

    ' Return error code
    Load = Err.Number
    
End Function

' Redirect: Redirect the link pointing to one word to another word
' Trust: Notifies if caller attempted to redirect a link which doesn't exist
' Arguments: lWordToUnlinkObjNo, lWordToSubObjNo
' Returns: Object number of word substituted
Public Function Redirect(lWordToUnlinkObjNo As Long, lWordToSubObjNo As Long) As Long

    ' Find the Link to change
    Dim objLinkToUnlink As clsLink
    Set objLinkToUnlink = mcolLinks(GenWordKey(lWordToUnlinkObjNo))
    
    ' Sanity check - alert if the link doesn't exist and exit the function
    If objLinkToUnlink Is Nothing Then
        MsgBox "clsLinks:Redirect attempted to unlink a non existant link!"
        Redirect = 0
        Exit Function
    End If
    
    ' Keep the link occurences
    Dim lTempLinkOccs As Long
    lTempLinkOccs = objLinkToUnlink.lOccs
    
    ' Relink to the word to substitute
    Remove objLinkToUnlink.lLinkNo
    Dim objNewRedirectedLink As clsLink
    Set objNewRedirectedLink = Add(lWordToSubObjNo, lTempLinkOccs)
    Redirect = objNewRedirectedLink.lLinkNo
    
End Function

' Remove: Remove a link from the collection
' Trust: Not removing a link such that the number of links will be <= 0
' Arguments: lLinkToRemoveObjNo
' Returns: Number of links removed excluding control words
Public Function Remove(lWordToRemoveObjNo As Long) As Long

    Dim objLinkToRemove As clsLink
    
    ' See if a link exists to the given word
    ' # Set objLinkToRemove = Find(lWordToRemoveObjNo)
    Set objLinkToRemove = mcolLinks(GenWordKey(lWordToRemoveObjNo))
        
    ' Exit if no such link exists
    If objLinkToRemove Is Nothing Then Remove = 0: Exit Function
    
    ' Return the number of occurences deleted
    Remove = objLinkToRemove.lOccs
    
    ' Update the internal statistics
    mlOccs = mlOccs - objLinkToRemove.lOccs
    
    ' Remove the link
    mcolLinks.Remove GenWordKey(objLinkToRemove.lLinkNo)

End Function

' Reweight: Reweight a specified link in this collection
' Trust: HIGH, doesn't check if link actually belongs to this collection
' Arguments: objLinkToReweight, lNewWeight
' Returns: clsLink reweighted
Public Function Reweight(objLinkToReweight As clsLink, lNewWeight As Long) As clsLink
    Dim lOldWeight As Long

    ' Sanity - make sure the new weight is positive
    If lNewWeight <= 0 Then MsgBox "ERROR: clsLinks:Reweight called with " + lNewWeight + ".  Should be +ve", vbExclamation: Exit Function
    
    With objLinkToReweight
    
        ' Keep old occurences on link object
        lOldWeight = .lOccs
    
        ' Adjust link object
        .lOccs = lNewWeight
    
        ' Adjust summary statistics on this collection, subtract old and add new
        mlOccs = mlOccs - lOldWeight + lNewWeight
        
    End With
    
    ' Return link that has been reweighted
    Set Reweight = objLinkToReweight
    
End Function

' Save: Invokes the save routine on all links in the collection and any other information
' Arguments: nFileNo - file number to use for save
' Returns: non-zero if error
Public Function Save(nFileNo As Integer) As Long

    ' Invoke save method on each link in the collection
    Dim objLink As clsLink
    For Each objLink In mcolLinks
        Save = objLink.Save(nFileNo): If Save Then Exit Function
    Next
    
    ' Signal everything went okay
    Save = 0
    
End Function

' --------------------------------------------------------------------------------
' ********************************************************************************
' --------------------------------------------------------------------------------
' CountWeight: Finds the total weight of links contained in the collection
' Trust: HIGH
' Arguments: NONE
' Returns: long
' *** BUT WHY IS THIS HERE?  SURELY WE ARE COMMITTED TO KEEPING A VALID COUNT ALREADY?***
Private Function CountWeight() As Long
    Dim objLink As clsLink
    CountWeight = 0
    
    ' Iterate through each link and add its weight
    For Each objLink In mcolLinks
        With objLink
            CountWeight = CountWeight + .lOccs
        End With
    Next
End Function

' --------------------------------------------------------------------------------
' ********************************************************************************
' --------------------------------------------------------------------------------

Public Property Get lOccs() As Long: lOccs = mlOccs: End Property
Public Property Let lOccs(lOccs As Long): mlOccs = lOccs: End Property

Private Sub Class_Initialize()

    ' Setup and initialise the collection
    Set mcolLinks = New clsCollection
    mcolLinks.Initialise vbObject
    
    mlOccs = 0
    
End Sub
    
Private Sub Class_Terminate()
    ' Debug.Print "clsLinks terminated"
End Sub
