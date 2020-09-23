VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   4965
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   6180
   HelpContextID   =   1200
   Icon            =   "Options.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame Frame2 
         Caption         =   "File Options"
         Height          =   1335
         Left            =   240
         TabIndex        =   16
         Top             =   1440
         Width           =   4935
         Begin VB.OptionButton radAppend 
            Caption         =   "Append new log to old log"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   720
            Width           =   2535
         End
         Begin VB.OptionButton radReplace 
            Caption         =   "Replace old log when recording"
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   360
            Value           =   -1  'True
            Width           =   3015
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Logging"
         Height          =   915
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   4935
         Begin VB.CheckBox chkAutoLog 
            Caption         =   "Automatically log conversations if logfile selected"
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   360
            Width           =   4335
         End
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Speech"
         Height          =   3345
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   5175
         Begin VB.CheckBox chkSpeakOnLoad 
            Caption         =   "Say a sentence on loading"
            Height          =   375
            Left            =   240
            TabIndex        =   8
            Top             =   360
            Width           =   2295
         End
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   0
      Left            =   240
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample1 
         Caption         =   "Reporting Level"
         Height          =   2265
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1695
         Begin VB.OptionButton radRepNone 
            Caption         =   "None"
            Height          =   375
            Left            =   240
            TabIndex        =   12
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton radRepHigh 
            Caption         =   "High"
            Height          =   375
            Left            =   240
            TabIndex        =   11
            Top             =   1440
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton radRepMed 
            Caption         =   "Medium"
            Height          =   375
            Left            =   240
            TabIndex        =   10
            Top             =   1080
            Width           =   1095
         End
         Begin VB.OptionButton radRepLow 
            Caption         =   "Low"
            Height          =   375
            Left            =   240
            TabIndex        =   9
            Top             =   720
            Width           =   1095
         End
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2490
      TabIndex        =   1
      Top             =   4455
      Width           =   1095
   End
   Begin ComctlLib.TabStrip tbsOptions 
      Height          =   4245
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7488
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Analysis"
            Key             =   "tabAnalysis"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Options for managing the way Actors analyze sentences"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Load"
            Key             =   "tabLoad"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Set Options for Loading Actors"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Logging"
            Key             =   "tabLogging"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Set options for logging conversations"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmOptions"
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

Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    
    ' Invoke context help
    SendKeys "{F1}"
    
End Sub

Private Sub cmdOK_Click()

    ' Set up new values in structures
    gobjOptions.gbSpeakOnLoad = chkSpeakOnLoad.Value
    
    gobjOptions.geAppendLog = IIf(radReplace, Overwrite, Append)
    
    gobjOptions.gbAutoLog = chkAutoLog.Value
    
    If radRepNone Then
        gobjOptions.meRepLevel = None
    ElseIf radRepLow Then gobjOptions.meRepLevel = Low
    ElseIf radRepMed Then gobjOptions.meRepLevel = Medium
    ElseIf radRepHigh Then gobjOptions.meRepLevel = High
    End If

    Unload Me
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    'handle ctrl+tab to move to the next tab
    If Shift = vbCtrlMask And KeyCode = vbKeyTab Then
        i = tbsOptions.SelectedItem.Index
        If i = tbsOptions.Tabs.Count Then
            'last tab so we need to wrap to tab 1
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
        Else
            'increment the tab
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i + 1)
        End If
    End If
End Sub

Private Sub Form_Load()
    
    ' Set up the data in the options form
    If gobjOptions.gbSpeakOnLoad Then chkSpeakOnLoad.Value = 1 Else chkSpeakOnLoad.Value = 0
    If gobjOptions.gbAutoLog Then chkAutoLog.Value = 1 Else chkAutoLog.Value = 0
    radReplace = IIf(gobjOptions.geAppendLog = Overwrite, True, False)
    radAppend = IIf(gobjOptions.geAppendLog = Append, True, False)

    ' Set up the report level information
    Select Case gobjOptions.meRepLevel
    Case None
        radRepNone = True
    Case Low
        frmOptions.radRepLow = True
    Case Medium
        frmOptions.radRepMed = True
    Case High
        frmOptions.radRepHigh = True
    Case Else
        Stop  ' Reporting has to be one of these four in the enumeration!
    End Select
    
    'center the form
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub

Private Sub tbsOptions_Click()
    
    Dim i As Integer
    'show and enable the selected tab's controls
    'and hide and disable all others
    For i = 0 To tbsOptions.Tabs.Count - 1
        If i = tbsOptions.SelectedItem.Index - 1 Then
            picOptions(i).Left = 210
            picOptions(i).Enabled = True
        Else
            picOptions(i).Left = -20000
            picOptions(i).Enabled = False
        End If
    Next
    
End Sub
