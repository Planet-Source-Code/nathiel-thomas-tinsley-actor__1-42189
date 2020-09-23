VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actor"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   1755
   ClientWidth     =   6735
   HelpContextID   =   1000
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   6735
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   10
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "New"
            Description     =   "New Actor"
            Object.ToolTipText     =   "Start a new actor"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Open"
            Description     =   "Open Actor"
            Object.ToolTipText     =   "Revive a previous actor"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Save"
            Description     =   "Save Actor"
            Object.ToolTipText     =   "Save the current actor"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "btnViewVocab"
            Object.ToolTipText     =   "Toggle the vocabulary view"
            Object.Tag             =   ""
            ImageIndex      =   4
            Style           =   1
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "btnViewMonitor"
            Object.ToolTipText     =   "Toggle the monitor view"
            Object.Tag             =   ""
            ImageIndex      =   5
            Style           =   1
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "btnRecord"
            Object.ToolTipText     =   "Log the current conversation"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "btnStop"
            Object.Tag             =   "Stop logging the conversation"
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Say 
      Caption         =   "Say"
      Default         =   -1  'True
      Height          =   735
      Left            =   7440
      TabIndex        =   4
      Top             =   5520
      Width           =   855
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   7035
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   2196
            MinWidth        =   2187
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   3254
            MinWidth        =   3263
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   3254
            MinWidth        =   3263
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   165
      Top             =   1965
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.TextBox Output 
      Height          =   4845
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   810
      Width           =   6015
   End
   Begin VB.TextBox User 
      Height          =   735
      Left            =   330
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   6060
      Width           =   6015
   End
   Begin VB.Label lblInput 
      Caption         =   "Input Area"
      Height          =   240
      Left            =   360
      TabIndex        =   6
      Top             =   5820
      Width           =   1575
   End
   Begin VB.Label lblConversation 
      Caption         =   "Conversation Area"
      Height          =   210
      Left            =   360
      TabIndex        =   5
      Top             =   570
      Width           =   1485
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   8
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":030A
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":041C
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":052E
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":0640
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":095A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":0C74
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":0F8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":12A8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu Voc 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu Vocab_New 
         Caption         =   "&New"
      End
      Begin VB.Menu Vocab_Open 
         Caption         =   "&Open"
      End
      Begin VB.Menu Vocab_Save 
         Caption         =   "&Save"
      End
      Begin VB.Menu Vocab_Save_As 
         Caption         =   "Save &As"
      End
      Begin VB.Menu Exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu View 
      Caption         =   "&View"
      Index           =   1
      Begin VB.Menu View_Vocabulary 
         Caption         =   "&Vocabulary"
      End
      Begin VB.Menu mnuView_Monitor 
         Caption         =   "&Monitor"
      End
   End
   Begin VB.Menu mnuActor 
      Caption         =   "&Actor"
      Index           =   3
      Begin VB.Menu mnuActor_Properties 
         Caption         =   "&Properties"
      End
   End
   Begin VB.Menu Tools 
      Caption         =   "&Tools"
      Index           =   3
      Begin VB.Menu Import_Text 
         Caption         =   "&Import Text"
      End
      Begin VB.Menu mnuTools_Options 
         Caption         =   "&Options"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "&Help"
      Index           =   10
      Begin VB.Menu mnuHelp_Contents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu About 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
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

' requires (incomplete list): clsLog


Option Explicit

Private Enum meActor  ' For passing to function to indicate who is speaking the string
    Person = 0
    Niall = 1
End Enum

Private mobjActorLog As clsLog

Private mstrActorFileTitle As String        ' Keep a record of the vocab file currently being worked on (path and file)
Private mstrExclFileTitle As String         ' Keep a record of the name of the last file used as a contextual exclusion list
Private mstrImportFileTitle As String       ' Keep a record of the file name of the last imported file
Private mstrActorPath As String
Private mstrExclPath As String
Private mstrImportPath As String


' UpdateStatus: Update all the information in the main window (inc the status panel and the actor name)
' Trust: LOW
' Returns: NONE
Public Sub UpdateStatus()

    ' Indicate in the titlebar the name of this program, the name of this actor
    ' and an * if it's changed since the last save
    frmMain.Caption = conMainCapPrefix & gobjCurrActor.strName
    frmMain.Caption = frmMain.Caption & IIf(gobjCurrActor.bChangedFlag, "*", "")
    
    UpdateStatusPanel

End Sub

' ********************************************************************************
' --------------------------------------------------------------------------------
' ********************************************************************************

' GetDashes: Generate a row of dashes as a 'tearoff'
' Trust: N/A
' Arguments: NONE
' Returns: string of dashes
Private Function GetDashes() As String

    ' Create the dashes
    Dim i As Integer
    For i = 1 To conOutputDashes: GetDashes = GetDashes & conDash: Next
    
End Function

' LoadActor: This function loads the Actor and carries out housekeeping tasks
' Trust:
' Arguments: strActorFileName, objActor
' Returns: non-zero if error
Private Function LoadActor(strActorFileName As String, objActor As clsActor) As Long

    Dim nFileNo As Integer
    Dim lReturnCode&
    
    On Error GoTo ErrorHandler
    
    ' Set up a new actor
    Set objActor = New clsActor
    objActor.Initialise gcolExclusions
        
    ' Assume for now that Vocab.Load is successful and the file is not bad
    mstrActorFileTitle = GetFileTitle(strActorFileName)
    
    ' Store the new(or unchanged) actor path
    mstrActorPath = GetPath(strActorFileName)
    
    ' Get a free file number and use it to open the selected file for binary reading and then close it
    nFileNo = FreeFile
    Open strActorFileName For Binary As nFileNo
    lReturnCode = objActor.Load(nFileNo)
    Close nFileNo
    
    ' Error checking
    If lReturnCode Then Err.Raise lReturnCode
        
    ' Perform necessary display updates on successfully opening this new file
    UpdateStatus
    
    ' Automatically start logging if option set and there is a valid logfile
    If gobjOptions.gbAutoLog And objActor.strLogFileTitle <> "" Then StartLog objActor.strLogPath & objActor.strLogFileTitle
    
    ' If set to, make the actor speak a sentence when its loaded, before the user has said anything.
    If gobjOptions.gbSpeakOnLoad Then UpdateOutput Niall, objActor.Reply()
    
    ' Return success
    LoadActor = 0
    
    Exit Function
    
ErrorHandler:

    Dim strMsg$
    
    Select Case Err.Number

    Case errLoadFileHigher
        strMsg = mstrActorFileTitle + " could not be loaded. It was created by an Actor of version " _
        + CStr(gnVersion) + ", but this program is version " + _
        CStr(GetVersion(App.Major, App.Minor, App.Revision)) + vbCrLf
        strMsg = strMsg + "Please look at Help->About to find out where to get the latest version of this program."
        
    Case Else
        strMsg = mstrActorFileTitle + " could not be loaded.  The file may be corrupted or not a type " + APPNAME + " file." + Chr(10) + Chr(10)
        strMsg = strMsg + "See logfile for more details."
        
    End Select
    
    MsgBox strMsg, vbCritical, "Loading Error"
    
    ' Remove the failure
    RmActor objActor
    
    ' Write out to the buffer so it can inspected without leaving the program
    gobjLog.FlushBuffer
    
    LoadActor = Err.Number
    
End Function

' LoadExclusions: Attempt to open the file which contains the contextual exclusions
' Trust:
' Arguments:
' Returns: bool - false success, true failure
Private Function LoadExclusions() As Boolean

    On Error GoTo ErrorHandler

    Dim nFileNo%, lReturnCode&
    
    ' Get a free file number and use it to open the selected file for ascii reading and then close it
    nFileNo = FreeFile
    Open mstrExclPath + mstrExclFileTitle For Input As nFileNo
    lReturnCode = gcolExclusions.Load(nFileNo): If lReturnCode Then Err.Raise lReturnCode
    Close nFileNo
    
    ' ### TEMPORARY ###
    gcolExclusions.Dbug
    
    LoadExclusions = False

    Exit Function
    
ErrorHandler:

    ' Notify but carry on
    Dim strMsg$
    strMsg = "Couldn't open " + mstrExclPath + mstrExclFileTitle + " to parse as an excluded words file!"
    MsgBox strMsg, vbExclamation
    gobjLog.Log strMsg, Warn
    
    LoadExclusions = True
    
End Function

' LoadSettings: Get initial startup settings or substitute defaults
' Trust: HIGH, assumes we will get settings back in the order we put them in
' Arguments: NONE
' Returns: NONE
Private Sub LoadSettings()

    On Error GoTo ErrorHandler

    Dim vRetSetting As Variant

    ' setup options information
    ' ### Need defaults to stop an error being generated if the value doesn't exist ###
    gobjOptions.meRepLevel = GetSetting(conAppName, conOptionsSection, "ReportingLevel", conDefRepLevel)
    gobjOptions.geAppendLog = IIf(GetSetting(conAppName, conOptionsSection, "AppendLog", conDefAppendLog), Append, Overwrite)
    gobjOptions.gbAutoLog = GetSetting(conAppName, conOptionsSection, "AutoLog", conDefAutoLog)
    gobjOptions.gbSpeakOnLoad = GetSetting(conAppName, conOptionsSection, "SpeakOnLoad", conDefSpeakOnLoad)
    
    ' setup windows information
    ' frmMain settings
    vRetSetting = GetSetting(conAppName, conWindowsSection, "Form1Left", conDefForm1Left)
    ' check validity
    If vRetSetting >= -frmMain.Width And vRetSetting <= Screen.Width Then frmMain.Left = vRetSetting
    vRetSetting = GetSetting(conAppName, conWindowsSection, "Form1Top", conDefForm1Top)
    ' check validity
    If vRetSetting >= -frmMain.Height And vRetSetting <= Screen.Height Then frmMain.Top = vRetSetting
    
    ' frmMonitor settings
    frmMonitor.Visible = GetSetting(conAppName, conWindowsSection, "frmMonitorVisible", conDefMonVisible)
    vRetSetting = GetSetting(conAppName, conWindowsSection, "frmMonitorLeft", conDefMonLeft)
    ' check validity
    If vRetSetting >= -frmMonitor.Width And vRetSetting <= Screen.Width Then frmMonitor.Left = vRetSetting
    vRetSetting = GetSetting(conAppName, conWindowsSection, "frmMonitorTop", conDefMonTop)
    ' check validity
    If vRetSetting >= -frmMonitor.Height And vRetSetting <= Screen.Height Then frmMonitor.Top = vRetSetting
    
    ' VocabForm settings
    frmVocab.Visible = GetSetting(conAppName, conWindowsSection, "VocabFormVisible", conDefVocVisible)
    vRetSetting = GetSetting(conAppName, conWindowsSection, "VocabFormLeft", conDefVocLeft)
    ' check validity
    If vRetSetting >= -frmVocab.Width And vRetSetting <= Screen.Width Then frmVocab.Left = vRetSetting
    vRetSetting = GetSetting(conAppName, conWindowsSection, "VocabFormTop", conDefVocTop)
    ' check validity
    If vRetSetting >= -frmVocab.Height And vRetSetting <= Screen.Height Then frmVocab.Top = vRetSetting

    ' path information
    mstrActorPath = GetSetting(conAppName, conPathsSection, "ActorPath")
    ' mstrExclPath = GetSetting(conAppName, conPathsSection, "ExclPath", App.Path + "\" + conDefExclPath)
    ' *** TEMPORARY choosing where the exclude.txt file is hasn't yet been implemented ***
    mstrExclPath = App.Path + "\" + conDefExclPath
    mstrImportPath = GetSetting(conAppName, conPathsSection, "ImportPath")
    gobjOptions.gstrDefLogPath = GetSetting(conAppName, conPathsSection, "LogPath")
    
    ' Files information
    mstrExclFileTitle = GetSetting(conAppName, conFilesSection, "ExclFileTitle", conDefExclFileTitle)
    
Exit Sub

ErrorHandler:
    MsgBox "Error loading registry settings, please report", vbExclamation
    
End Sub

' NewActor: Delete all traces of the old actor and set up a blank new one
' Trust:
' Arguments: NONE
' Returns: a blank Actor
Private Sub NewActor(objActor As clsActor)
    
    ' Set up a new actor
    Set objActor = New clsActor
    objActor.Initialise gcolExclusions
    
    ' update status information
    UpdateStatus
    
    ' Signal new actor on output screen
    PrintDashes
    
End Sub

' PrintDashes: Print a row of dashes in the output text box
' Trust: LOW
' Returns: NONE
' Arguments: NONE
Private Sub PrintDashes()

    ' Create the tearoff in memory
    Dim strTearOff As String
    strTearOff = GetDashes() & vbCrLf

    ' Print
    Output.SelStart = Len(Output.Text)
    Output.SelText = strTearOff
    Output.SelStart = Len(Output.Text)
    
End Sub

' QueryCloseActor: Query the user as to whether they want to save an actor
' Trust:
' Arugments: NONE
' Returns: bool - TRUE if user has chosen to cancel the operation, FALSE if user has saved or approved
Private Function QueryCloseActor() As Boolean

    ' sanity - if the actor hasn't changed don't bother with the query
    If Not gobjCurrActor.bChangedFlag Then QueryCloseActor = False: Exit Function

    ' If the actor has changed since the last save then query the user if they want to save or abort exit
    Dim strQueryMsg As String, nStatus As Integer
    strQueryMsg = gobjCurrActor.strName & " has changed since the last save.  Do you want to save the changes?"
        
    nStatus = MsgBox(strQueryMsg, vbExclamation Or vbYesNoCancel, gobjCurrActor.strName)
    
    Select Case nStatus
    ' User does want to save but also still wants to exit
    Case vbYes
        Vocab_Save_Click
        QueryCloseActor = False
    ' user approves of exit
    Case vbNo
        QueryCloseActor = False
    ' User does not want to exit after all
    Case vbCancel
        QueryCloseActor = True
    End Select
    
End Function

' QueryLogFile: Asks the user which file to use as a log file
' Trust:
' Arguments: NONE
' Returns: string - empty if no file was selected
Private Function QueryLogFile() As String

    On Error GoTo ErrorHandler  ' ENABLE ERRORHANDLER to catch dialog box errors and file allocation errors
    
    ' Get the text file to use as a log (or append to)
    CommonDialog1.Flags = cdlOFNHideReadOnly
    CommonDialog1.DialogTitle = "Choose Log File"
    CommonDialog1.Filter = TEXTFILEDESC + "|" + TEXTFILEFILT + "|" + ALLFILEDESC + "|" + ALLFILEFILT
    
    ' If this is a new actor (that is, no log path has been set yet) then use the last one
    CommonDialog1.InitDir = IIf(gobjCurrActor.strLogPath = "", gobjOptions.gstrDefLogPath, gobjCurrActor.strLogPath)
    CommonDialog1.filename = IIf(gobjCurrActor.strLogFileTitle = "", gobjOptions.gstrDefLogTitle, gobjCurrActor.strLogFileTitle)
    
    CommonDialog1.ShowOpen
    
    ' IF THE CANCEL BUTTON IS PRESSED ON THE COMMON DIALOG BOX, AN ERROR IS GENERATED WHICH HAS TO BE TRAPPED!
    
    ' return the complete file name
    QueryLogFile = CommonDialog1.filename

    Exit Function
    
ErrorHandler:
    Select Case Err.Number
        ' Trap the cancel error raised if cancel is clicked in the dialog box
        Case cdlCancel
            ' Do nothing
    End Select
    
End Function

' RmActor: Remove the given actor, tidying up displays
' Trust:
Private Sub RmActor(objActor As clsActor)

    ' close off any logfile for the old actor
    StopLog
    
    ' Make sure save button now invokes save as by deleting any existing save file title
    mstrActorFileTitle = ""
    
    ' Remove the existing vocabulary view list
    frmVocab.DeleteList
    
    ' Need to teardown to remove references.  Class_Terminate won't call if refs still in place
    objActor.TearDown
    
End Sub

' SaveActor: Save the actor
' Trust:
' Arguments: strFileName - Complete filename to use for save
' Returns: non-zero if error
Private Function SaveActor(strFileName As String) As Long

    On Error GoTo ErrorHandler
    
    Dim nSaveFileNo As Integer
    Dim lReturn&
    
    ' Sanity - if the file name given already exists and is read-only then signal and bum out
    If Dir(strFileName, (vbNormal Or vbHidden Or vbSystem)) <> "" Then
        If (GetAttr(strFileName) And vbReadOnly) Then SaveActor = vbObjectError + errReadOnly: Exit Function
    End If
    
    ' set up file for saving
    nSaveFileNo = FreeFile
    gobjLog.Log "<start save>", Info  ' ***DEBUG LINE
    Open strFileName For Binary As nSaveFileNo  ' Open using the next free file number
    
    ' If the file title is different from the actor's name and the actor's name hasn't been changed
    ' Then assume the file name is a proper name, and subsitute it for the Actor's default name
    If gobjCurrActor.strName = conDefName Then
        Dim strFileTitle As String
        strFileTitle = GetFileTitle(strFileName)
        gobjCurrActor.strName = Left$(strFileTitle, FindLast(strFileTitle, conActorFileSuffix) - 1)
    End If
    
    ' let main actor object do its stuff
    lReturn = gobjCurrActor.Save(nSaveFileNo): If lReturn Then gobjLog.Log "Could not save Actor: " + gobjCurrActor.strName + " to " + strFileName, Error: Err.Raise lReturn
    
    ' close off the file
    Close nSaveFileNo
    gobjLog.Log "<end save>", Info
    SaveActor = True
    
    ' Update display status
    UpdateStatus
    
    ' Signal everything went okay
    SaveActor = 0
    
    Exit Function

ErrorHandler:
    MsgBox "An error occurred while trying to save " + gobjCurrActor.strName + "!  See logfile for more details.", vbExclamation, "Error"
    SaveActor = Err.Number
    
End Function

' SaveSettings: Save settings in the registry
' Trust:
' Arguments: NONE
' Returns: NONE
Private Sub SaveSettings()

    ' Save options settings to the registry
    SaveSetting conAppName, conOptionsSection, "SpeakOnLoad", gobjOptions.gbSpeakOnLoad
    SaveSetting conAppName, conOptionsSection, "AppendLog", gobjOptions.geAppendLog
    SaveSetting conAppName, conOptionsSection, "AutoLog", gobjOptions.gbAutoLog
    SaveSetting conAppName, conOptionsSection, "ReportingLevel", gobjOptions.meRepLevel
    
    ' Save windows settings to the registry
    SaveSetting conAppName, conWindowsSection, "Form1Left", frmMain.Left
    SaveSetting conAppName, conWindowsSection, "Form1Top", frmMain.Top
    SaveSetting conAppName, conWindowsSection, "frmMonitorVisible", frmMonitor.Visible
    SaveSetting conAppName, conWindowsSection, "frmMonitorLeft", frmMonitor.Left
    SaveSetting conAppName, conWindowsSection, "frmMonitorTop", frmMonitor.Top
    SaveSetting conAppName, conWindowsSection, "VocabFormVisible", frmVocab.Visible
    SaveSetting conAppName, conWindowsSection, "VocabFormLeft", frmVocab.Left
    SaveSetting conAppName, conWindowsSection, "VocabFormTop", frmVocab.Top
    
    ' Save path settings to the registry
    SaveSetting conAppName, conPathsSection, "ActorPath", mstrActorPath
    SaveSetting conAppName, conPathsSection, "ExclPath", mstrExclPath
    SaveSetting conAppName, conPathsSection, "ImportPath", mstrImportPath
    SaveSetting conAppName, conPathsSection, "LogPath", gobjOptions.gstrDefLogPath
    
    ' Save file settings to registry
    SaveSetting conAppName, conFilesSection, "ExclFileTitle", mstrExclFileTitle

End Sub

' StartLog: Get a log file from the user and start logging the conversation
' Trust:
' Arguments: strLogFileName
' Returns: NONE
Private Sub StartLog(strLogFileName$)

    On Error GoTo ErrorHandler
    
    Dim lReturnCode&
    
    lReturnCode = gobjCurrActor.mobjLog.Oublier(strLogFileName, gobjOptions.geAppendLog): If lReturnCode Then Err.Raise lReturnCode
    
    ' update toolbar
    Toolbar1.Buttons.Item(conRecordButtonKey).Enabled = False
    Toolbar1.Buttons.Item(conStopButtonKey).Enabled = True
    
    ' update monitor
    With gobjCurrActor
        frmMonitor.Output "Started logging conversation to " & .strLogPath & .strLogFileTitle, Low
    End With
    
    Exit Sub
    
ErrorHandler:
    
    Select Case Err.Number
        ' Trap the cancel error raised if cancel is clicked in the dialog box and do nothing
        Case cdlCancel

        Case errNoFilename
            MsgBox "Tried to frmMain:StartLog without a logfile name!", vbExclamation, "Error"
            
        ' Read only error
        Case vbObjectError + errReadOnly
            MsgBox GetFileTitle(strLogFileName) & errReadOnlyLogPrompt, vbExclamation, errReadOnlyLogTitle
            
    End Select
    
End Sub

' StopLog: Stop a conversation log that has already been started
Private Sub StopLog()

    gobjCurrActor.mobjLog.Ferme
    
    ' update toolbar
    Toolbar1.Buttons.Item(conStopButtonKey).Enabled = False
    Toolbar1.Buttons.Item(conRecordButtonKey).Enabled = True
    
    ' update monitor
    With gobjCurrActor
        frmMonitor.Output "Stopped logging conversation to " & .strLogPath & .strLogFileTitle, Low
    End With
    
End Sub

' UpdateLog: Update the logfile with the given text
' Trust: NONE
' Arguments: strToLog
' Returns: NONE
Private Sub UpdateLog(strToLog As String)

    gobjCurrActor.mobjLog.Log strToLog, Gen
    
End Sub

' UpdateOutput: Insert a new string on the end of the output box and scroll it down
Private Sub UpdateOutput(ByVal a As meActor, ByVal s As String)
    
    ' Decide on appropriate prefix to sentence
    Dim prefix As String
    If a = Niall Then prefix = gobjCurrActor.strName & ": " Else prefix = "User: "
    
    ' Insert sentence and then a Chr$(13) & Chr$(10) at the end
    s = prefix & s
    
    ' ***HANDLE SCROLLING***
    ' Basically, this code sets the insert point (or 'selection') to the end of the current output text
    ' Then sets the selection text to the string, which inserts the text at the insertion point
    ' And moves the insertion point to the new end of output.  Beautiful, simple scrolling.
    Output.SelStart = Len(Output.Text)
    Output.SelText = s & vbCrLf
    Output.SelStart = Len(Output.Text)
    
    ' update logfile
    UpdateLog s
    
End Sub

' UpdateStatusPanel: Quick and dirty way of updating the little status panels with word and links info
Private Sub UpdateStatusPanel()

    StatusBar1.Panels(1).Text = "Words: " + CStr(gobjCurrActor.colVocab.lWordsCount)
    StatusBar1.Panels(2).Text = "Ordinary Links: " + CStr(gobjCurrActor.colVocab.lOrdOccsCount)
    
    ' Depending on whether contextual linking has been selected, display contextual links or 'N/A' as appropriate
    Dim strConLinks As String
    strConLinks = IIf(gobjCurrActor.bContextual, CStr(gobjCurrActor.colVocab.lConOccsCount), "N/A")

    StatusBar1.Panels(3).Text = "Contextual Links: " + strConLinks
    
End Sub

' ********************************************************************************
' --------------------------------------------------------------------------------
' ********************************************************************************

Private Sub Form_Initialize()

    ' Initialize variables
    mstrActorFileTitle = ""
    Set mobjActorLog = New clsLog
    
    ' Initial jobs
    Randomize
    
    ' Create global objects
    ' Apparantly, creating the object when the class its in is initialized results in less overhed
    ' than if the object were declared as new
    Set gcolExclusions = New clsExclusions
    ' Set gobjCurrActor = New clsActor
    ' gobjCurrActor.Initialise gcolExclusions
    Set gobjOptions = New clsOptions
    Set gobjLog = New clsLog
    
    ' Start program logging
    gobjLog.Oublier App.Path + "\" + LOGFILENAME, Overwrite, True
    
    ' Attempt to load up registry settings
    LoadSettings
    
    ' Attempt to load a contextual exclusions file
    LoadExclusions
    
    ' Press any buttons on the toolbar that need pressing
    If frmMonitor.Visible Then Toolbar1.Buttons.Item(conMonitorButtonKey).Value = tbrPressed
    If frmVocab.Visible Then Toolbar1.Buttons.Item(conVocabButtonKey).Value = tbrPressed
    
    ' Initialize help file
    App.HelpFile = App.Path & "\" & conHelpFile
    
    ' If an actor file has been passed as a cmd line parameter (from Explorer, say) then parse it
    Dim strArgs$: strArgs = Command()
    
    ' If some argument has been passed then assume it is the filename of an actor, and attempt to load
    If Len(strArgs) Then
        ' If loading wasn't successful substitute in a blank actor
        If LoadActor(RemoveQuotes(strArgs), gobjCurrActor) Then NewActor gobjCurrActor
    Else
        ' If no initial argument was given we have to set up a new actor anyway
        NewActor gobjCurrActor
    End If

End Sub

' Form_QueryUnload: Checks to see if the actor has changed since the last save and queries if so
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    ' If the actor has changed since the last save then query the user if they want to save or abort exit
    If QueryCloseActor() = True Then Cancel = True Else Cancel = False
    
End Sub

' Form_Unload: Perform environment tasks and unload objects
Private Sub Form_Unload(Cancel As Integer)

    ' Close off any logfiles (program and actor) still open
    StopLog
    gobjLog.Ferme
    
    ' Save environmental settings to the registry
    SaveSettings

    ' Unload each form in the program
    Dim f As Form
    For Each f In Forms: Unload f: Next
    
End Sub

Private Sub mnuActor_Properties_Click()

    ' Set up the data in the properties form
    frmProperties.txtName = gobjCurrActor.strName
    frmProperties.txtMinConSize = gobjCurrActor.nMinConSize
    frmProperties.chkContext = IIf(gobjCurrActor.bContextual, 1, 0)
    frmProperties.chkExclude = IIf(gobjCurrActor.bExclude, 1, 0)
    frmProperties.chkLimReply = IIf(gobjCurrActor.bLimReplyLength, 1, 0)
    frmProperties.txtPrefMaxReply = gobjCurrActor.nPrefMaxReplyLength
    frmProperties.cmbThreshold = gobjCurrActor.fTermProportionReq * 100

    ' Enable/Disable the correct controls
    frmProperties.chkLimReply_Click
    
    ' Display the properties form ready for changes
    frmProperties.Show vbModal, Me
    
End Sub

Private Sub mnuHelp_Contents_Click()

    ' Show the help dialog box (API function)
    WinHelp hWnd, App.Path & "\" & conHelpFile, conHelpTab, 0
    
End Sub

Private Sub mnuTools_Options_Click()
    
    ' Display the options form ready for changes
    frmOptions.Show vbModal, Me
    
End Sub

Private Sub mnuView_Monitor_Click()

    ' Display the monitor window
    frmMonitor.Show
    
End Sub

' Say_Click: This is the really the main loop of the program
' Say_Click: Parses contents of user's textbox
' Say_Click: Say button is also clicked by default when user presses return in the user textbox
Private Sub Say_Click()

    ' Update the conversation window to reflect the user's input
    UpdateOutput Person, User.Text

    ' Analyze the user's input sentence by sentence
    ' We have to transfer the user string to another variable because User.text appears to be passed only by value
    ' Even though I've specified by reference, so it never gets chopped down to nothing to trip the while condition
    Dim strUserInput As String: strUserInput = User.Text
    While strUserInput <> ""
        gobjCurrActor.Analyze AnalyzeSentence(strUserInput)
    Wend

    ' Update the status display
    UpdateStatus

    ' Generate a reply to the user
    UpdateOutput Niall, gobjCurrActor.Reply()
    
    User.Text = ""
    
End Sub

Private Sub About_Click()
    frmAbout.Show
End Sub

' Exit_Click: Basically do the same thing as if exit was selected from the control menu
Private Sub Exit_Click()
    Unload Me
End Sub

' Toolbar button routine
Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)

    Dim strFileName$

    Select Case Button.key
    Case "New"
    
        Vocab_New_Click
        
    Case "Open"
    
        Vocab_Open_Click
        
    Case "Save"
    
        Vocab_Save_Click
        
    Case conVocabButtonKey
    
        ' Toggle vocabulary view status according to the status of the button on the toolbar
        ' By the time this executed the button state has been changed by the click
        If frmMain.Toolbar1.Buttons.Item(conVocabButtonKey).Value = tbrPressed Then
            View_Vocabulary_Click
        Else
            frmVocab.Hide
        End If
        
    Case conMonitorButtonKey
    
        ' Toggle Monitor view status
        If frmMain.Toolbar1.Buttons.Item(conMonitorButtonKey).Value = tbrPressed Then
            mnuView_Monitor_Click
        Else
            frmMonitor.Hide
        End If
        
    Case conRecordButtonKey
    
        ' Start logging
        strFileName = QueryLogFile()
        If strFileName <> "" Then StartLog strFileName
        
    Case conStopButtonKey
    
        ' Stop logging
        StopLog

    End Select
    
End Sub

Private Sub View_Vocabulary_Click()

    ' Display the vocabulary form
    frmVocab.Show

End Sub

' Vocab_New_Click: Delete a new clsActor and initialise a new blank one
Private Sub Vocab_New_Click()

    ' check what the user wants to do with the current actor if it has been modified
    If QueryCloseActor() = True Then Exit Sub
    
    ' Delete the old clsActor first, to prevent any bugs with unremoved references, collections, etc.
    RmActor gobjCurrActor
    
    ' Set up the new Actor
    NewActor gobjCurrActor
    
End Sub

Private Sub Vocab_Open_Click()

    ' check what the user wants to do with the current actor if it has been modified
    If QueryCloseActor() = True Then Exit Sub

    On Error GoTo ErrorHandler  ' ENABLE ERRORHANDLER
    
    ' Delete the old clsActor first, to prevent any bugs with unremoved references, collections, etc.
    RmActor gobjCurrActor
    
    ' IF THE CANCEL BUTTON IS PRESSED ON THE COMMON DIALOG BOX, AN ERROR IS GENERATED WHICH HAS TO BE TRAPPED!
    CommonDialog1.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
    CommonDialog1.DialogTitle = "Open Actor"
    CommonDialog1.Filter = NIALLFILEDESC + "|" + "*" + conActorFileSuffix + "|" + ALLFILEDESC + "|" + ALLFILEFILT
    CommonDialog1.InitDir = mstrActorPath
    CommonDialog1.filename = mstrActorFileTitle
    CommonDialog1.ShowOpen
    
    ' If loading wasn't successful substitute in a blank actor
    If LoadActor(CommonDialog1.filename, gobjCurrActor) Then NewActor gobjCurrActor
    
    Exit Sub
    
ErrorHandler:
    Select Case Err.Number
        ' Trap the cancel error raised if cancel is clicked in the dialog box
        Case cdlCancel
            ' Do nothing

    End Select
End Sub

Private Sub Vocab_Save_Click()

    On Error GoTo ErrorHandler
    
    Dim lReturn&
    
    ' Invoke the save as routine if this file hasn't been saved before
    If mstrActorFileTitle = "" Then
        Vocab_Save_As_Click
    Else
        lReturn = SaveActor(mstrActorPath & mstrActorFileTitle): If lReturn Then Err.Raise lReturn
    End If
    
    Exit Sub
    
ErrorHandler:
    Select Case Err.Number
    
        ' Trap the cancel error raised if cancel is clicked in the dialog box
        Case cdlCancel
            ' Do nothing
            
        ' Read only error
        Case vbObjectError + errReadOnly
            MsgBox mstrActorFileTitle & errReadOnlyActorPrompt, vbExclamation, errReadOnlyActorTitle
            
        Case Else
            ' The user has been notified of other errors in SaveActor(), so do nothing here
        
    End Select
    
End Sub

' Save Niall's current vocabulary to a valid external file
Private Sub Vocab_Save_As_Click()

    On Error GoTo ErrorHandler
    
    Dim lReturn&
    
    With CommonDialog1
        .Flags = cdlOFNPathMustExist Or cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
        .DialogTitle = "Save Actor As"
        .Filter = NIALLFILEDESC + "|" + "*" + conActorFileSuffix + "|" + ALLFILEDESC + "|" + ALLFILEFILT
        .InitDir = mstrActorPath
    
        ' If there is no pre-existing file for this actor, then construct one using the Actor's name.
        If mstrActorFileTitle <> "" Then
            .filename = mstrActorFileTitle
        Else
            .filename = gobjCurrActor.strName + conActorFileSuffix
        End If
        
        .ShowSave
    
        ' Capture the save_as file title for future reference
        mstrActorFileTitle = .FileTitle
        
        ' Capture the path
        mstrActorPath = Left$(.filename, Len(.filename) - Len(mstrActorFileTitle))
            
        lReturn = SaveActor(mstrActorPath & mstrActorFileTitle): If lReturn Then Err.Raise lReturn
    
    End With
    
    Exit Sub
    
ErrorHandler:
    Select Case Err.Number
        ' Trap the cancel error raised if cancel is clicked in the dialog box
        Case cdlCancel
            ' Do nothing
            
        ' Read only error
        Case vbObjectError + errReadOnly
            MsgBox mstrActorFileTitle & errReadOnlyActorPrompt, vbExclamation, errReadOnlyActorTitle
            
        Case Else
            ' Anything else should have been dealt with by SaveActor(), so do nothing
            
    End Select
    
End Sub

Private Sub Import_Text_Click()

    On Error GoTo ErrorHandler
    
    ' Get the text file to import
    CommonDialog1.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
    CommonDialog1.DialogTitle = "Import"
    CommonDialog1.Filter = TEXTFILEDESC + "|" + TEXTFILEFILT + "|" + ALLFILEDESC + "|" + ALLFILEFILT
    CommonDialog1.InitDir = mstrImportPath
    CommonDialog1.filename = mstrImportFileTitle
    CommonDialog1.ShowOpen
    
    ' Capture selected path and filename
    mstrImportFileTitle = CommonDialog1.FileTitle
    mstrImportPath = Left$(CommonDialog1.filename, Len(CommonDialog1.filename) - Len(mstrImportFileTitle))
    
    Dim fileno As Integer: fileno = FreeFile
    gobjLog.Log "<start import>", Info   ' ***DEBUG LINE
    
    Open mstrImportPath & mstrImportFileTitle For Input As fileno
    
    ' Debug.Print "Opened:"; mstrImportPath & mstrImportFileTitle  ' ***DEBUG LINE
    
    ' Do the stuff
    gobjCurrActor.Import fileno

    ' Finish with the file
    Close fileno
    
    ' Perform necessary display updates on successfully opening this new file
    UpdateStatus
    
    Exit Sub
    
ErrorHandler:
    Select Case Err.Number
        ' Trap the cancel error raised if cancel is clicked in the dialog box
        Case cdlCancel
            ' Do nothing
            
        ' The overflow error will be thrown if we try to parse a text file with neither Windows or Unix delimiters
        ' And should occur if any one sentence is too big for the string datatype (I think!)
        Case cdlOVERFLOW
            MsgBox "No appropriate sentence breaks in the text file!  Import aborted.", vbOKOnly Or vbCritical, "Warning!"
    End Select
    
End Sub
