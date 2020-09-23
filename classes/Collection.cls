VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCollection"
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

' clsCollection: A base class built on the collection datatype
' clsCollection: Implements ability to find items.
' Trust: Assumes item keys are unique.
' Trust: Assumes all items are of the same type

Option Explicit

Private Const mconInvalidType = -1  ' Constant to signal that an item is not the same type as those already in the collection

Private mbIsInit As Boolean
Private mcolItems As Collection
Private mnVarType As Integer        ' Stores the variable type of items based on the first item added

' Add: Add a new item to the collection using a supplied key
' Trust:
' Arguments: varData, lKey, lBefore - indicate whether the item should be entered at a certain position in the collection
' Returns: integer - 0 if success, 1 if key already exists, 2 if object is the wrong type (define these elsewhere)
Public Function Add(ByRef varData As Variant, Optional ByVal strKey As String = "", Optional ByVal lBefore As Long = 0) As Integer

    ' Sanity - collection needs to be initialised before doing ANYTHING
    If Not mbIsInit Then Stop

    Dim varSearch As Variant
    
    ' Signal nonsense error until we know otherwise
    Add = -1
    
    ' Sanity - check a key and data has been given
    If mnVarType = vbObject Then
        If varData Is Nothing Then Exit Function
    Else
        If IsNull(varData) Or IsEmpty(varData) Then Exit Function
    End If
        
    ' Sanity - Otherwise, only accept the item if it is the same type as the first item
    If VarType(varData) <> mnVarType Then Add = 2: Exit Function
    
    ' Different behaviour is required depending on whether we are dealing with objects or non-objects
    If mnVarType = vbObject Then
    
        ' See if the key already exists if a key has been specified
        If strKey <> "" Then Set varSearch = Item(strKey) Else Set varSearch = Nothing
        
        ' If not, then add to the collection
        If varSearch Is Nothing Then
        
            ' Pass through the initial arguments given for a normal collection add
            If strKey <> "" Then
                If lBefore <> 0 Then
                    mcolItems.Add varData, strKey, lBefore
                Else
                    mcolItems.Add varData, strKey
                End If
            Else
                If lBefore <> 0 Then
                    mcolItems.Add varData, , lBefore
                Else
                    mcolItems.Add varData
                End If
            End If
            
            ' Signal success
            Add = 0
            
        ' Otherwise, signal
        Else
            Add = 1
        End If
    
    Else
    
        ' See if the key already exists
        varSearch = Item(strKey)

        ' If not, then add to the collection
        If IsEmpty(varSearch) Then
            mcolItems.Add varData, strKey
            Add = 0
            
        ' Otherwise, signal
        Else
            Add = 1
        End If
    
    End If
    
End Function

' Count: Returns number of items
' Trust:
' Arguments: NONE
' Returns: long - number of items
Public Function Count() As Long

    Count = mcolItems.Count
    
End Function

' Initialise: Needs to be called to explicitly tell the collection what type of objects it will be receiving
' Trust: Assumes nObjectType is a sensible value
' Arguments: nObjectType
' Returns: NONE
Public Function Initialise(ByVal nObjectType As Integer)

    ' Sanity - can't be sensible if negative
    If nObjectType < 0 Then Exit Function
    
    ' Set up
    mnVarType = nObjectType
    
    ' Signal
    mbIsInit = True
    
End Function

' Item: Search this collection to see if an item specified by a given index string already exists
' Trust: Assumes mnVarType has already been defined
' Arguments: strKey
' Returns: variant - item if there's a match, Empty or nothing if there is no match.
Public Function Item(ByRef varKey As Variant) As Variant
Attribute Item.VB_UserMemId = 0

    ' Sanity - collection needs to be initialised before doing ANYTHING
    If Not mbIsInit Then Stop

' --- Here, I'm using a dodgy method to return nothing if no link matching the link number is found ---
' --- Basically, a collection returns an 'invalid procedure' error if you try to retrieve on a non-existent key ---
' --- However, I'm intercepting this error with the handler and returning nothing instead, which is the behaviour I want ---
    On Error GoTo ErrorHandler

    If mnVarType = vbObject Then
        ' Try to retrieve a valid item from the collection
        Set Item = mcolItems(varKey)
    Else
        Item = mcolItems(varKey)
    End If

Exit Function

ErrorHandler:
    ' If nothing has been found, this error occurs, which I can intercept and return nothing to signal no match
    If Err.Number = 5 Or Err.Number = 91 Then
    
        ' Set find dependent on whether we're dealing with objects or non-objects
        If mnVarType = vbObject Then
            Set Item = Nothing
        Else
            Item = Empty
        End If
    
    End If

End Function

' NewEnum must return the IUnknown interface of a
' collection's enumerator.
' (This copied from Visual Basic Books Online to allow 'for each ... in' to be used)
' Bit of a hack really - needs Tools - Procedure Attributes to have ID=-4 and this procedure hidden
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mcolItems.[_NewEnum]
End Function

' Remove: Remove an item from the index
' Trust: Assumes varIndex is valid(!)
' Arguments: varIndex
' Returns NONE
Public Function Remove(varIndex As Variant)

    mcolItems.Remove (varIndex)

End Function

Private Sub Class_Initialize()

    ' Load var type with invalid value
    mnVarType = mconInvalidType
    mbIsInit = False
    
    ' Create collection
    Set mcolItems = New Collection
    
End Sub
