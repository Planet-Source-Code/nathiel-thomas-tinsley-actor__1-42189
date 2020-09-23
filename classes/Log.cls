VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLog"
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

' clsLog: This is a general class for performing all the functions and holdng variables associated with using a logfile
' requires: basGeneral

Option Explicit

Public Enum eLogModes                           ' These enum figures needed for legacy compatibility
    Overwrite = 0
    Append = -1
End Enum

Public Enum eEventType
    Error = vbLogEventTypeError
    Warn = vbLogEventTypeWarning
    Info = vbLogEventTypeInformation
    Gen = &H8
End Enum
    
Private Const OFFSET& = vbObjectError + &H200   ' Don't ask - apparantly we have to offset the offset by 512
Private Const WRITETHRESHOLD& = 5000            ' When the number of unwritten lines reaches this point close & reopen the logfile to free memory.
Private Const errNoFilename& = OFFSET + 1       ' Error code return if no filename is given to :oublier.
Private Const errLogAlready& = OFFSET + 2       ' Error code return if :oublier is called when logging already active
Private Const errNoLog& = OFFSET + 3            ' Error code return if ferme is closed when no logfile exists

Public mbDebug As Boolean                       ' If set, tee log entries out to debug
Public mbLogStatus As Boolean                   ' Are we logging at the moment?
Public meAppend As eLogModes
Private mlLinesToWrite As Long                  ' Number of lines which VB has not yet physically written to the logfile
Private mnLogFileNo As Integer
Public mstrLogTitle As String                   ' Keep a record of the name of the last file used to append a log
Public mstrLogPath As String

' Ferme: Close an open logfile
' Trust:
' Arguments: NONE
' Returns: non-zero if error
Public Function Ferme() As Long

    On Error GoTo ErrorHandler
    
    ' Sanity - signal error if we close when there is no file
    If Not mbLogStatus Then Err.Raise errNoLog
    
    ' Handle file manipulations
    Close mnLogFileNo
    
    ' update status
    mbLogStatus = False
    
    Exit Function
    
ErrorHandler:
    
    Ferme = Err.Number

End Function

' FlushBuffer: Write out existing contents of buffer by closing and repoening
' Trust:
' Arguments:
' Returns: non-zero if error
Public Function FlushBuffer() As Long

    On Error GoTo ErrorHandler

    Close mnLogFileNo
    mnLogFileNo = FreeFile
    Open mstrLogPath + mstrLogTitle For Append As #mnLogFileNo
    
    ' Reset stats
    mlLinesToWrite = 0
    
    Exit Function
    
ErrorHandler:
    
    FlushBuffer = Err.Number
    
End Function

' Log: Echo the given input to the logfile
' Trust:
' Arguments: strEntry
' Returns: non-zero if error
Public Function Log(strEntry$, eType As eEventType) As Long

    On Error GoTo ErrorHandler
    
    Dim strPrefix$, strOutput$
    
    ' Sanity - if no logfile exists, return as such
    If Not mbLogStatus Then Err.Raise errNoLog
    
    ' Generate appropriate prefix
    Select Case eType
        Case Error
            strPrefix = "ERROR: "
        Case Warn
            strPrefix = "Warn: "
        Case Info
            strPrefix = "Info: "
        Case Else
            strPrefix = ""
    End Select
            
    ' Make log entry and tee to debug if desired
    strOutput = strPrefix + strEntry
    Print #mnLogFileNo, strOutput
    If mbDebug Then Debug.Print strOutput
    
    ' Upgrade internal stats
    mlLinesToWrite = mlLinesToWrite + 1
    
    ' If we've reached the limit, flush the output buffer by closing and reopening the file
    If mlLinesToWrite >= WRITETHRESHOLD Then FlushBuffer
    
    ' Signal no error
    Log = 0
    
    Exit Function
    
ErrorHandler:
    
    Log = Err.Number
    
End Function

' Oublier: Open the filename given as a logfile
' Trust:
' Arguments: strFileName, eMode - default is append, bDebug - tee all log entries out to debug if true
' Returns: non-zero if error
Public Function Oublier(Optional strFileName$, Optional eMode As eLogModes = Append, Optional bDebug As Boolean) As Long

    On Error GoTo ErrorHandler
    
    ' Sanity - If we're already logging then this should not be executed
    If mbLogStatus Then Err.Raise errLogAlready
    
    ' Sanity - if the log file name given already exists and is read-only then hurl
    If strFileName <> "" Then
        If Dir(strFileName, (vbNormal Or vbHidden Or vbSystem)) <> "" Then
            If (GetAttr(strFileName) And vbReadOnly) Then Err.Raise vbObjectError + errReadOnly
        End If
    End If
    
    ' Store the new filename information if given, otherwise, keep what we already have
    If strFileName = "" Then
        ' Only carry on if we have an existing file title if no filename has been passed
        If mstrLogTitle = "" Then Err.Raise errNoFilename
    Else
        mstrLogTitle = GetFileTitle(strFileName)
        mstrLogPath = GetPath(strFileName)
    End If
    
    ' Start logging, replacing or appending to any existing file as appropraite
    mnLogFileNo = FreeFile
    If eMode = Append Then
        Open strFileName For Append Shared As #mnLogFileNo
    Else
        Open strFileName For Output Shared As #mnLogFileNo
    End If
    meAppend = eMode
    
    ' Signal that logging has now started
    mbLogStatus = True
    
    mbDebug = bDebug
    
    ' Signal no error
    Oublier = 0
    
    Exit Function
    
ErrorHandler:
    
    Oublier = Err.Number
    
End Function

' --------------------------------------------------------------------------------
' ********************************************************************************
' --------------------------------------------------------------------------------


' --------------------------------------------------------------------------------
' ********************************************************************************
' --------------------------------------------------------------------------------

Private Sub Class_Initialize()

    mbDebug = False
    mstrLogTitle = ""
    mstrLogPath = ""
    mbLogStatus = False
    mnLogFileNo = -1
    mlLinesToWrite = 0
    
End Sub
