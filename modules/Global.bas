Attribute VB_Name = "basGlobal"
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

Option Explicit

' Declare an Actor and an Options for use by the whole program
Public gobjCurrActor As clsActor                                ' This is just a reference to gobjCast.mobjCurrActor to improve readability
Public gcolExclusions As clsExclusions                          ' Holds words to be excluded from contextualisation
Public gobjOptions As clsOptions                                ' Log file for the program itself
Public gobjLog As clsLog                                        ' Program log file

' Global variables
Public gnVersion As Integer                                     ' Temp for version number.  Used during loading

' Global ENUMs
' order indicates what type of link is requested
' Needs to correspond with the index numbers of the option buttons on the vocabulary display
Public Enum order
    backward = 0
    forward = 1
    con = 2
End Enum
' eRepLevel enumerates reporting levels
Public Enum geRepLevel
    None = 0
    Low = 1  ' Not yet used
    Medium = 2
    High = 3
End Enum
    

' GLOBAL CONSTANTS

    ' Visual Basic constants - in help files but not defined by VB itself!
    Public Const vbLogAuto = &H0                                    ' Writes to the appropriate log depending on the OS
    Public Const vbLogOverwrite = &H10

    ' Registry (here because conAppName needs to be defined early)
    Public Const conAppName = "Actor"
    Public Const conFilesSection = "Files"
    Public Const conOptionsSection = "Options"
    Public Const conPathsSection = "Paths"
    Public Const conWindowsSection = "Windows"
    
    ' Program (here because other constants dependant)
    Public Const APPNAME$ = "Actor"                                 ' Use in dialog boxes rather than hard-coding app name
    Public Const EMAIL$ = "mersault@users.sourceforge.net"          ' Contact details for bug reports, etc.
    
    ' Actor
    Public Const conDefName As String = "Anonymous"

    ' Analysis
    Public Const conDefaultMinConSize = 5
    
    ' Contextual
    Public Const conDefExclude = True
    Public Const conUpperConSetWeight = 5000                       ' Settable upper weighting for contextual links
    
    ' Dialog boxes
    Public Const conActorFileSuffix As String = ".atr"
    Public Const NIALLFILEDESC As String = APPNAME + " (*" + conActorFileSuffix + ")" ' Historically actor was called Niall in development
    Public Const TEXTFILEFILT As String = "*.txt"
    Public Const TEXTFILEDESC As String = "Text Files (" + TEXTFILEFILT + ")"
    Public Const ALLFILEFILT As String = "*.*"
    Public Const ALLFILEDESC As String = "All Files (" + ALLFILEFILT + ")"
    
    ' Display
    Public Const conMainCapPrefix As String = conAppName & " - "
    Public Const conDash As String = "-"                           ' Character to use as a dash
    Public Const conOutputDashes As Integer = 92                   ' Number of dashes to create a 'tearoff' on the output form textbox
    
    ' Error codes
    Public Const OFFSET& = vbObjectError + &H200                    ' VB Online book says we need to offset the offset
    Public Const errReadOnly As Integer = 1                         ' For read only file errors
    Public Const errReadOnlyActorPrompt As String = " is set to read-only.  Please choose another file or clear the read-only attribute."
    Public Const errReadOnlyLogPrompt As String = " is set to read-only.  To continue logging, please clear the read-only attribute and press record or choose another log file."
    Public Const errReadOnlyActorTitle As String = "Actor file is set to read-only!"
    Public Const errReadOnlyLogTitle As String = "LogFile is set to read-only!"
    Public Const errNoFilename& = OFFSET + 1                        ' Error code return if no filename is given to :oublier.
    Public Const errLogAlready& = OFFSET + 2                        ' Error code return if :oublier is called when logging already active
    Public Const errNoLog& = OFFSET + 3                             ' Error code return if ferme is closed when no logfile exists
    Public Const errMagicMisplaced As Long = OFFSET + &H80          ' On loading, the magic byte isn't where it should be
    Public Const errLoadFileHigher& = OFFSET + &H100                ' Load file version is higher than this program!
    Public Const cdlOVERFLOW As Integer = 6
    
    ' Files
    Public Const conDefExclPath = "rules\"                           ' These are defined here rather than placed by installation in the registry for source compilation convenience
    Public Const conDefExclFileTitle = "exclude.txt"
    Public Const LOGFILENAME$ = APPNAME + ".log"                    ' Program logfile
    Public Const MAGICBYTE As Byte = &H66                           ' Embed in files for error checking
    
    ' Help codes
    Public Const conHelpFile = "Actor.hlp"                          ' Actor's help file which should be in same directory
    
    ' Logging
    Public Const conDefAutoLog As Boolean = False
    Public Const conDefAppendLog = Append              ' Only append logs if the user asks for it
    
    ' Monitor
    Public Const conDefRepLevel As Integer = 3                      ' Set default reporting level to High
    
    ' Responses
    Public Const conDefBackwardLinking As Boolean = True            ' Default
    Public Const conDefLimReplyLength As Boolean = True             ' Actor seems to make a bit more sense if we curtail jabber
    Public Const conDefPrefMaxReplyLength As Integer = 6            ' Default preferred length for replies when we are curtailing jabber
    Public Const conDefTermProportionReq As Single = 0.1            ' Default proportion for actor to consider word a valid terminator when curtailing
    Public Const conDefSpeakOnLoad As Boolean = True                ' Set default speak on load option
    
    ' Toolbar buttons
    Public Const conVocabButtonKey$ = "btnViewVocab"                ' Key for Vocab View button on toolbar
    Public Const conMonitorButtonKey$ = "btnViewMonitor"
    Public Const conLogButtonKey$ = "btnLog"
    Public Const conRecordButtonKey$ = "btnRecord"
    Public Const conStopButtonKey As String = "btnStop"
    
    ' Windows
    Public Const conDefMonVisible As Boolean = False
    Public Const conDefVocVisible As Boolean = False
    Public Const conDefForm1Left As Integer = 0                    ' Set default window positions
    Public Const conDefForm1Top As Integer = 1140
    Public Const conDefMonLeft As Integer = 6825
    Public Const conDefMonTop As Integer = 6255
    Public Const conDefVocLeft As Integer = 6825
    Public Const conDefVocTop As Integer = 1140
    Public Const conInstColWidth As Integer = 590                  ' Width of the instances column in vocablistview
    Public Const conConColWidth As Integer = 590                   ' Width of the contextual column in vocablistview
    
    ' Unclassified (must classify!)
    Public Const conFirstWordObjNo As Long = 1                     ' The first word takes this number.
                                                                   ' CANNOT be zero since some functions use this to signal no word
    Public Const conEOSObjNo As Long = -1                          ' Word object number for the sentence ending 'word'                                                        ' Needs to be lower than the first word object number
    Public Const conSOSObjNo As Long = -2                          ' Word object number for the sentence starting 'word'
    Public Const conEOSString As String = "<< Sentence End >>"     ' Text string for the sentence ending 'word'
    Public Const conSOSString As String = "<< Sentence Start >>"   ' Text string for the sentence starting 'word'
    Public Const conWordKey As String = "Word"                     ' Key to use for word column in collections
    Public Const conPrevKey As String = "PrevLinks"                ' Key to use for previous links column
    Public Const conNextKey As String = "NextLinks"                ' Key to use for next links column
    Public Const conConKey As String = "Contextual"                ' Key to use for contextual column
