Attribute VB_Name = "basHelp"
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

' *** Declarations to WinHelp API functions and another function to hide details
' *** Pinched extensively from http://www.vbexplorer.com/winhelpapi.asp
' *** February 2001
Option Explicit

' Constants
Public Const conHelpTab = &HF  ' Setting to show help tab in help dialog box


' API Declarations
Declare Function WinHelp Lib "user32" Alias "WinHelpA" _
        (ByVal hWnd As Long, _
         ByVal lpHelpFile As String, _
         ByVal wCommand As Long, _
         ByVal dwData As Long) As Long
Declare Function WinHelpTopic Lib "user32" Alias "WinHelpA" _
         (ByVal hWnd As Long, _
          ByVal lpHelpFile As String, _
          ByVal wCommand As Long, _
          ByVal dwData As String) As Long


