VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOptions"
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

' clsOptions: This is the options class for storing all options flags, configuration information, etc
Option Explicit

' Logging
Public gstrDefLogTitle$
Public gstrDefLogPath$

' Signal whether a log should be appended to an existing file or replaced.
Public geAppendLog As eLogModes

' Should a log be automatically started on actor load if there is a valid logfile?
Public gbAutoLog As Boolean

' Flag whether the actor should say something when it is initially loaded
Public gbSpeakOnLoad As Boolean

' OTHER
Public meRepLevel As geRepLevel  ' Activity reporting level.  Range gnNone..gnHigh

' Set up the default settings, which will be invoked when this class is first created
Private Sub Class_Initialize()

    ' ### Defaults set up in frmMain:LoadSettings ###
    gstrDefLogTitle = ""
    gstrDefLogPath = ""
    geAppendLog = Append

End Sub
