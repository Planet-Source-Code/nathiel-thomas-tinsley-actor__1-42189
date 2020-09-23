VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLookUp"
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

' clsLookUp: contain look up data

Public mcolLookUp As clsCollection  ' Contains lookup data - public so we can later enumerate if necessary.

' Add: Add a new lookup
' Trust:
' Arguments: lObjNo, strWord
' Returns: int - FALSE if successful, 1 if lookup already exists, 2 if collection is wrong type.
Public Function Add(ByVal lObjNo As Long, ByRef strWord As String) As Integer

    ' Sanity - check a word has been passed
    If strWord = "" Then MsgBox "clsLookUp:Add - empty strWord passed!", vbCritical: Exit Function
    
    ' Add in new lookup
    Add = mcolLookUp.Add(strWord, CStr(lObjNo))
    
End Function

' Item: Lookup the word string signified by the given object number
' Trust:
' Arguments: lObjNo
' Returns: string - empty string signifies nothing found
Public Function Item(ByVal lObjNo As Long) As String
Attribute Item.VB_UserMemId = 0

    Dim strReturn As String
    
    strReturn = mcolLookUp(CStr(lObjNo))
    
    If IsEmpty(strReturn) Then Item = "" Else Item = strReturn
    
End Function

Private Sub Class_Initialize()

    Set mcolLookUp = New clsCollection
    mcolLookUp.Initialise vbString
    
End Sub
