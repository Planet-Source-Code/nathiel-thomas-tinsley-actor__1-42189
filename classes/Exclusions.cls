VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExclusions"
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

' clsExclusions: Use clsCollections as a based for handling the words to exclude from contextualisation

Public mcolExclusions As clsCollection  ' Only left public for iteration purposes

' Item: Lookup the word string signified by the given object number
' Trust:
' Arguments: strSearch
' Returns: string - empty string signifies nothing found
Public Function Item(ByRef strSearch As String) As String
Attribute Item.VB_UserMemId = 0

    Dim strReturn As String
    
    Item = mcolExclusions(strSearch)
    
End Function

' Load: Load attempts to parse an external file into a collection of words to be excluded from contextualisation
' Trust: The File Number passed is valid
' Arguments: nFileNo
' Returns: non-zero if error
Public Function Load(nFileNo As Integer) As Long

    On Error GoTo ErrorHandler

    Dim strImportLine$

    ' The contextual exclusion file consists of words seperated by newlines
    ' Import these words and store them in the collection
    While Not EOF(nFileNo)
        strImportLine = basGeneral.InputLine(nFileNo)
        mcolExclusions.Add strImportLine, strImportLine
    Wend
    
    ' Signal everything went a-okay
    Load = 0
    
    Exit Function
    
ErrorHandler:

    ' Return error number
    Load = Err.Number

End Function

' ********************************************************************************
' --------------------------------------------------------------------------------
' ********************************************************************************

Private Sub Class_Initialize()
    
    ' Make collection only accept strings
    Set mcolExclusions = New clsCollection
    mcolExclusions.Initialise vbString
    
End Sub

' ********************************************************************************
' --------------------------------------------------------------------------------
' ********************************************************************************

Public Sub Dbug()

    Dim v As Variant
    
'    For Each v In mcolExclusions
'        Debug.Print v
'    Next
    
    Debug.Print "Exclusion words loaded: " + CStr(mcolExclusions.Count)
        
End Sub
