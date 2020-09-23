VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form frmVocab 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Vocabulary"
   ClientHeight    =   4785
   ClientLeft      =   6870
   ClientTop       =   1425
   ClientWidth     =   7725
   HelpContextID   =   2000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   7725
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton objLinkOption 
      Caption         =   "Contextual"
      Height          =   375
      Index           =   2
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.OptionButton objLinkOption 
      Caption         =   "Previous"
      Height          =   375
      Index           =   0
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.OptionButton objLinkOption 
      Caption         =   "Next"
      Height          =   375
      Index           =   1
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Value           =   -1  'True
      Width           =   1215
   End
   Begin ComctlLib.ListView LinkListView 
      Height          =   3975
      Left            =   3960
      TabIndex        =   1
      Top             =   600
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   7011
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      SmallIcons      =   "VocabImageList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Linked word"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Occurences"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   ""
         Object.Width           =   2540
      EndProperty
   End
   Begin ComctlLib.ListView VocabListView 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   7011
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      SmallIcons      =   "VocabImageList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   "Word"
         Object.Tag             =   ""
         Text            =   "Word"
         Object.Width           =   2893
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   "PrevLinks"
         Object.Tag             =   ""
         Text            =   "Previous"
         Object.Width           =   1041
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   "NextLinks"
         Object.Tag             =   ""
         Text            =   "Next"
         Object.Width           =   1041
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   "Contextual"
         Object.Tag             =   ""
         Text            =   "Contextual"
         Object.Width           =   0
      EndProperty
   End
   Begin ComctlLib.ImageList VocabImageList 
      Left            =   840
      Top             =   -120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Vocab.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Vocab.frx":031A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuVocab 
      Caption         =   "&Vocab"
      Visible         =   0   'False
      Begin VB.Menu mnuVocab_Correct 
         Caption         =   "&Correct"
      End
      Begin VB.Menu mnuVocab_Delete 
         Caption         =   "&Delete"
      End
   End
   Begin VB.Menu mnuConLink 
      Caption         =   "ConLink"
      Visible         =   0   'False
      Begin VB.Menu mnuConLink_AdWeight 
         Caption         =   "&Adjust Weighting"
      End
   End
End
Attribute VB_Name = "frmVocab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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

' Provision of subsidary functions for maintaining the list view in terms of Words
Option Explicit

Dim meSelLinkOption As order  ' Currently selected link option

' AddWord: Add a word to the vocabulary list
' Trust: LOW
' Arguments: objWordToAdd
' Returns: NONE
Public Sub AddWord(objWordToAdd As clsWord)
    
    ' Add the word
    Dim objAddedItem As ListItem
    Set objAddedItem = VocabListView.ListItems.Add(, GenWordKey(objWordToAdd), objWordToAdd.strWord, , 1)  ' Hardcoded small icon number temporarily
    
    ' Add instances and contextual link information
    objAddedItem.SubItems(1) = CStr(objWordToAdd.colPrevLinks.lOccs)
    objAddedItem.SubItems(2) = CStr(objWordToAdd.colNextLinks.lOccs)
    objAddedItem.SubItems(3) = StringForListView(objWordToAdd.colConLinks.lOccs)
End Sub

' Remove all items in the existing vocab list
Public Sub DeleteList()
    VocabListView.ListItems.Clear
End Sub

' Remove an existing Word, assuming Word does exist
Public Sub Remove(ByRef n As clsWord)

    VocabListView.ListItems.Remove GenWordKey(n)
    
End Sub

' Update an existing Word
Public Sub Update(n As clsWord)

    Dim w As ListItem

    ' Retrieve the vocab list item for this word using its key
    Set w = VocabListView.ListItems(GenWordKey(n))
    Debug.Assert Not w Is Nothing
    
    ' Update number of links
    ' We only need to look at one link list to get the number of links, since both next and prev wil lbe identical
    w.SubItems(1) = CStr(n.colPrevLinks.lOccs)
    w.SubItems(2) = CStr(n.colNextLinks.lOccs)
    w.SubItems(3) = StringForListView(n.colConLinks.lOccs)
    
    ' If the word being updated is currently selected, then update the link list view as well
    If n Is DeGenWordKey(VocabListView.SelectedItem.key) Then VocabListView_Click
    
End Sub

' UpdateLink: Update a link in the link list view
' Arguments: lLinkToUpdateNo
' Returns: NONE
Public Sub UpdateLink(lLinkToUpdate As clsLink)

    Dim objLinkListItem As ListItem
    
    ' Retrieve the listitem for this link
    Set objLinkListItem = LinkListView.ListItems(GenWordKey(lLinkToUpdate.lLinkNo))
    ' Sanity check - Something has gone wrong if we try to update a link which isn't in the current list view
    Debug.Assert Not objLinkListItem Is Nothing
    
    ' Update weighting
    objLinkListItem.SubItems(1) = CStr(lLinkToUpdate.lOccs)
    
End Sub

' ********************************************************************************
' --------------------------------------------------------------------------------
' ********************************************************************************

' AddLink: Add the link to the link list view, either by creating a new entry or updating an existing link entry
' Trust: MEDIUM, assumes the link is valid, and doesn't already exist in the link view
' Arguments: objLinkToAdd
' Returns: NONE
Private Sub AddLink(objLinkToAdd As clsLink)

    ' On Error GoTo ErrorHandler

    Dim strLinkKey$
    Dim objAddedItem As ListItem, objFoundItem As ListItem

    ' Resolve word associated with our link
    Dim objLinkedWord As clsWord: Set objLinkedWord = gobjCurrActor.colVocab.Resolve(objLinkToAdd.lLinkNo)
    
    ' Sanity - If the link doesn't lookup to a valid word then write an error but continue
    If objLinkedWord Is Nothing Then
        gobjLog.Log "frmVocab:AddLink tried to add a non existant word with objno " + CStr(objLinkToAdd.lLinkNo) + "!", Warn
        Exit Sub
    End If

    ' Add the link to the link view
    strLinkKey = GenWordKey(objLinkedWord)
    Set objAddedItem = LinkListView.ListItems.Add(, strLinkKey, objLinkedWord.strWord, , 2)
    
    ' Update other link entry details
    objAddedItem.SubItems(1) = CStr(objLinkToAdd.lOccs)
    
    Exit Sub
    
ErrorHandler:
    Stop
    
End Sub

' DispLinks: Display the links of the selected word in the LinkListView, according to which mode has been selected
' Trust: HIGH - assumes objWordForLinks is a valid word
' Arguments: objWordForLinks
' Returns: NONE
Private Sub DispLinks(objWordForLinks As clsWord)

    ' Debug - Program should never pass DispLinks Nothing
    Debug.Assert Not objWordForLinks Is Nothing
    
    ' Sanity - If nothing is passed, don't attempt to do anything
    If objWordForLinks Is Nothing Then Exit Sub
    
    ' Display word links for the clicked word, depending on which link mode is selected
    Select Case meSelLinkOption
    ' Backward links
    Case backward
        DispLinksB objWordForLinks.colPrevLinks
    ' Forward links
    Case forward
        DispLinksB objWordForLinks.colNextLinks
    ' Contextual links
    Case con
        DispLinksB objWordForLinks.colConLinks
    End Select
    
End Sub

' DispLinksB: Should only be called from DispLinks.  Given a set of links, displays these in the list view
' Trust: HIGH, assumes link collection contains only links
' Arguments: objLinksToView
' Returns: NONE
Private Sub DispLinksB(objLinksToView As clsLinks)

    ' Delete whatever is already in the view
    LinkListView.ListItems.Clear
    
    ' Sanity check - has a links object been passed?
    If objLinksToView Is Nothing Then Stop
    
    ' Add each valid ordinary previous link to the link list view
    Dim objLink As clsLink, lLinkWord As clsWord
    ' # For Each objLink In objLinksToView
    For Each objLink In objLinksToView.mcolLinks
        AddLink objLink
    Next
    
End Sub

' StringForListView: Convert an integer into an appropriate string to be added to the listview
' Trust: N/A
' Arguments: lNumToConvert
' Returns: string
Private Function StringForListView(lNumToConvert As Long) As String

    ' If number is zero then use a -, otherwise use the number
    If lNumToConvert = 0 Then StringForListView = "-" Else StringForListView = CStr(lNumToConvert)
    
End Function

' ********************************************************************************
' --------------------------------------------------------------------------------
' ********************************************************************************

Private Sub Form_Initialize()
    meSelLinkOption = forward  ' Setup initially selected link view mode when vocbulary is first displayed
End Sub

' Form_QueryUnload: When user tries to close the vocab view, hide it rather than unload it
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    ' If it is the user closing the vocabulary view, then hide it instead
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        frmVocab.Hide
        
        ' Unpress the View Vocabulary button on the main window toolbar
        frmMain.Toolbar1.Buttons.Item(conVocabButtonKey).Value = tbrUnpressed
        
    End If
    
End Sub

' LinkListView_DblClick: If a word has been double clicked in the link list view,
' LinkListView_DblClick: then change the focus of the Vocablistview to point to this word and
' LinkListView_DblClick: change the LinkListView to point to its links.
Private Sub LinkListView_DblClick()

    ' Sanity - if no word is selected then don't do anything
    If LinkListView.SelectedItem Is Nothing Then Exit Sub
    
    ' Highlight the double clicked word on the VocabListView
    Set VocabListView.SelectedItem = VocabListView.ListItems(LinkListView.SelectedItem.key)
    VocabListView.SelectedItem.EnsureVisible

    ' Display word links for the clicked word, depending on which link mode is selected
    ' Equivalent to acting as if the user clicked this word themselves
    VocabListView_Click

End Sub

' LinkListView_MouseUp: When the right mouse button is clicked on the link list view, pop up the contextual link menu
Private Sub LinkListView_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' Check if right mouse button was clicked and option selected is contextual
    If Button = vbRightButton And meSelLinkOption = con Then PopupMenu mnuConLink  ' Display the contetual link menu as a pop-up menu.
    
End Sub

' mnuConLink_AdWeight_Click: Allow the user to adjust the contextual weighting of the selected word
' Trust:
' Arguments: NONE
' Returns: NONE
Private Sub mnuConLink_AdWeight_Click()

    Dim objLinkSelected As clsLink
    Dim objWordForLinks As clsWord
    
    ' Sanity - abort if there's no selected link
    If LinkListView.SelectedItem Is Nothing Then Exit Sub

    ' Retrieve the link clicked in the list view, first retrieving the word its associated with
    Set objWordForLinks = DeGenWordKey(VocabListView.SelectedItem.key)
    Set objLinkSelected = objWordForLinks.colConLinks.Find(KeyToObjNo(LinkListView.SelectedItem.key))
    
    ' Request the new weight from the user with a dialog box
    Dim strInput As String
    strInput = InputBox("Enter new weighting", "Adjust Contextual Weighting", CStr(objLinkSelected.lOccs))
    
    ' If OK, retrieve the new weight
    If strInput = "" Then Exit Sub  ' Cancel returns empty string
    
    ' Keep the new weight between one and an arbitrary constant upper bound
    Dim lNewValue As Long: lNewValue = CLng(strInput)
    If lNewValue < 1 Then lNewValue = 1
    If lNewValue > conUpperConSetWeight Then lNewValue = conUpperConSetWeight
    
    ' Set the appropriate contextual links entry to the new weighting
    gobjCurrActor.Reweight objWordForLinks, objLinkSelected, lNewValue

    ' Update the view
    UpdateLink objLinkSelected
    Update objWordForLinks
    
End Sub

' Vocab_Correct_Click:
Private Sub mnuVocab_Correct_Click()

    VocabListView.StartLabelEdit
    
End Sub

' Vocab_Delete_Click: Delete the selected word from the database
' Vocab_Delete_Click: Set the redirect flag to point to Nothing, signals EOS
' *** MIGHT BE NEATER TO HAVE AN EVENT RAISED HERE, RATHER THAN EMBEDDING CODE IN THE VOCABFORM OBJECT ***
Private Sub mnuVocab_Delete_Click()

    Dim colWordsToUpdate As clsWords, objWord As clsWord

    ' Retrieve the word clicked in the vocabulary view
    Dim objWordToDelete As clsWord
    Set objWordToDelete = DeGenWordKey(VocabListView.SelectedItem.key)
    
    ' Remove word from database
    Set colWordsToUpdate = gobjCurrActor.colVocab.Remove(objWordToDelete)
    
    ' Remove word from vocablist display
    Remove objWordToDelete
    
    ' Update the statistics of all affected words
    For Each objWord In colWordsToUpdate.mcolWords
        Update objWord
    Next
    
    ' Update main display word/occ counts
    frmMain.UpdateStatus
    
End Sub

Private Sub objLinkOption_Click(Index As Integer)

    ' Load the option status with the button now selected
    meSelLinkOption = Index
    
    ' Change the columns seen in the vocabulary list according to the option selected
    If meSelLinkOption = backward Or meSelLinkOption = forward Then
        ' Show only the word and the number of instances in the vocabulary view
        VocabListView.ColumnHeaders(conPrevKey).Width = conInstColWidth
        VocabListView.ColumnHeaders(conNextKey).Width = conInstColWidth
        VocabListView.ColumnHeaders(conConKey).Width = 0
    Else
        ' Show only the word and the number of contextual instances in the vocabulary view
        VocabListView.ColumnHeaders(conPrevKey).Width = 0
        VocabListView.ColumnHeaders(conNextKey).Width = 0
        VocabListView.ColumnHeaders(conConKey).Width = conConColWidth
    End If
    
    ' Update the link display, after checking a word has been selected
    If Not VocabListView.SelectedItem Is Nothing Then
        VocabListView_Click
    End If

End Sub

' ********************************************************************************
' --------------------------------------------------------------------------------
' ********************************************************************************

' Locate the Word and change accordingly
' If a word the same as NewString already exists then append the links of the old word to the existing word
' If no such word exists, leave the old Word intact and just change the word.
' *** MIGHT BE NEATER TO HAVE AN EVENT RAISED HERE, RATHER THAN EMBEDDING CODE IN THE VOCABFORM OBJECT ***
Private Sub VocabListView_AfterLabelEdit(Cancel As Integer, NewString As String)
    
    ' objSelectedNode - selected Word; objFoundNode - found Word
    Dim objSelectedNode As clsWord, objFoundNode As clsWord
    
    ' First, locate the position of the Word that's been selected
    Set objSelectedNode = DeGenWordKey(VocabListView.SelectedItem.key)
    
    ' Second, check if the newstring already exists as some other Word
    Set objFoundNode = gobjCurrActor.colVocab.Find(NewString)
    
    ' If the word does already exist, then copy over the links from the existing word
    ' and delete the existing word
    ' *** CAN'T DO THIS THE OTHER WAY AROUND BY DELETING THE SELECTED NODE ***
    ' *** CAUSES CRASH***
    If Not objFoundNode Is Nothing Then
    
        ' First, merge the words
        gobjCurrActor.colVocab.Merge objSelectedNode, objFoundNode
        
        ' Second, remove the merging source Word from the viewed list.  This has to be done second because other functions update it
        Remove objFoundNode
        
    End If
    
    ' Update main screen and vocablistview
    frmMain.UpdateStatus
    frmVocab.Update objSelectedNode
    
    ' If there is no existing Word for the new string, just change the old Word
    ' Update the selected node with the string its been changed to
    gobjCurrActor.colVocab.Edit objSelectedNode, NewString
    
End Sub

' VocabListView_Click: When a word is selected, display the appropriate collection of links depending on the option chosen
' Trust: N/A
' Arguments: NONE
' Returns: NONE
Private Sub VocabListView_Click()

    ' Sanity - if no vocabulary word has been selected then don't do anything
    If VocabListView.SelectedItem Is Nothing Then Exit Sub
    
    ' Retrieve the word clicked in the vocabulary view
    Dim objWordForLinks As clsWord
    Set objWordForLinks = DeGenWordKey(VocabListView.SelectedItem.key)
    
    ' Display word links for the clicked word, depending on which link mode is selected
    DispLinks objWordForLinks
    
End Sub

Private Sub VocabListView_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)

    ' NOT IMPLEMENTED BECAUSE SORTS OCCUR WITH NUMBERS AS STRINGS RATHER THAN NUMBERS.
    ' WILL NEED TO FIX THIS
    
    ' When a column header is clicked, change the view so contents are sorted by that header
    ' We want the alphabetical column to sort ascending but the numbers to be descending
'    VocabListView.SortKey = ColumnHeader.Index - 1 ' 0-based
'    ' If word column (first)
'    If VocabListView.SortKey = 0 Then
'        VocabListView.SortOrder = lvwAscending
'    ' Otherwise number column
'    Else
'        VocabListView.SortOrder = lvwDescending
'    End If
    
End Sub

' VocabListView_MouseUp: Pop up the contextual menu for vocabulary on a right mouse button click
Private Sub VocabListView_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' Check if right mouse button was clicked.
    If Button = vbRightButton Then
    
        ' Don't do anything unless word selected is a real word
        If IsRealWord(KeyToObjNo(VocabListView.SelectedItem.key)) Then PopupMenu mnuVocab
        
    End If
    
End Sub
