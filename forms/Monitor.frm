VERSION 5.00
Begin VB.Form frmMonitor 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Monitor"
   ClientHeight    =   4575
   ClientLeft      =   6870
   ClientTop       =   6540
   ClientWidth     =   4215
   HelpContextID   =   3000
   Icon            =   "Monitor.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtOutput 
      Height          =   4335
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmMonitor"
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

' Output: Output a string if the given reporting level meets the current threshold
' Trust:
' Arguments: strToOutput, eRepLevel
' Returns: NONE
Public Sub Output(strToOutput As String, eRepLevel As geRepLevel)

    ' Output a string if the given reporting level meets the current threshold
    If eRepLevel <= gobjOptions.meRepLevel Then PrintString strToOutput

End Sub

' ********************************************************************************
' --------------------------------------------------------------------------------
' ********************************************************************************

' PrintString: Insert a new string on the end of the output box and scroll it down
Private Sub PrintString(strLatestOutput As String)
    
    ' Insert a Chr$(13) & Chr$(10) on the end of the string
    strLatestOutput = strLatestOutput & vbCrLf
    
    ' ***HANDLE SCROLLING***
    ' Basically, this code sets the insert point (or 'selection') to the end of the current output text
    ' Then sets the selection text to the string, which inserts the text at the insertion point
    ' And moves the insertion point to the new end of output.  Beautiful, simple scrolling.
    txtOutput.SelStart = Len(txtOutput.Text)
    txtOutput.SelText = strLatestOutput
    txtOutput.SelStart = Len(txtOutput.Text)
    
End Sub

' Form_QueryUnload: When user tries to close the vocab view, hide it rather than unload it
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    ' If it is the user closing the vocabulary view, then hide it instead
    If UnloadMode = vbFormControlMenu Then
    
        Cancel = True
        Me.Hide
        
        ' Unpress the View Vocabulary button on the main window toolbar
        frmMain.Toolbar1.Buttons.Item(conMonitorButtonKey).Value = tbrUnpressed
        
    End If
    
End Sub
