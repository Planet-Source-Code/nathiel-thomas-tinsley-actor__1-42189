VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLink"
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

' clsLink: This contains linking information to connect words
Option Explicit

' Private mcolParent As clsLinks
Private mlLinkNo As Long        ' Object number of linked word.  Write once
Private mlOccs As Long          ' Number of times this link has occured
Private mbIsInit As Boolean     ' Flag to signal whether this link has already had its details set

' Create: Create the initial data of this link
' Trust: N/A
' Arguments: lLinkNo - write once linked node number
' Arguments: lOccs - initial number of occurences
' Returns: clsLink - reference to this object
Public Function Create(lLinkNo As Long, lOccs As Long) As clsLink

    ' Sainty - Check this link hasn't already been initialised
    Debug.Assert mbIsInit = False
    
    ' Fill in the blanks
    ' Set mcolParent = colParent
    mlLinkNo = lLinkNo
    mlOccs = lOccs
    mbIsInit = True
    
    ' Update parent statistics
 '   mcolParent.lOccs = mcolParent.lOccs + mlOccs
    
    ' Return this link (handy for calling functions)
    Set Create = Me
End Function

' Save: Saves all the information in the link to the given file number
' Arguments: nFileNo - File number to use for saving
' Returns: non-zero if error
Public Function Save(nFileNo As Integer) As Long

    On Error GoTo ErrorHandler
    
    ' Save the important link details
    Put nFileNo, , mlLinkNo
    Put nFileNo, , mlOccs
    
    Save = False ' Don't need to report any errors
    
    Exit Function
    
ErrorHandler:
    Save = Err.Number
    
End Function

' lLinkNo: Read only except by :create
' Trust: NONE
' Arguments: NONE
' Returns: Number of next link word
Property Get lLinkNo() As Long: lLinkNo = mlLinkNo: End Property

' lOccs: Read-write number of occurences on this link
Property Get lOccs() As Long: lOccs = mlOccs: End Property

Property Let lOccs(lOccs As Long)

    ' Sanity - make sure this link has been created before doing anything
    If mbIsInit = False Then MsgBox "ERROR: Attempt to set clsLink.lOccs before the link has been initialised!", vbExclamation: Exit Property
    
    ' Sanity - Make sure the value being set won't make link occurences zero or negative
    If lOccs <= 0 Then MsgBox "ERROR: Attempt to set clsLink.lOccs with " + lOccs + ", a non positive value!, vbExclamation"
    
    ' Set the variable
    mlOccs = lOccs
    
End Property

Private Sub Class_Initialize()
'    Debug.Print "CALLED clsLink:Class_Initialize"
    ' Signal that the contents of this class haven't yet been set
    mbIsInit = False
End Sub
