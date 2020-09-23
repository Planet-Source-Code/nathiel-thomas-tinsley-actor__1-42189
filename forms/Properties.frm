VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form frmProperties 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Actor Properties"
   ClientHeight    =   4920
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   6150
   HelpContextID   =   1100
   Icon            =   "Properties.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame Frame1 
         Caption         =   "Reply Length"
         Height          =   1860
         Left            =   255
         TabIndex        =   15
         Top             =   1605
         Width           =   5160
         Begin VB.ComboBox cmbThreshold 
            Height          =   315
            ItemData        =   "Properties.frx":000C
            Left            =   600
            List            =   "Properties.frx":0031
            TabIndex        =   18
            Text            =   "Replace Me"
            Top             =   1155
            Width           =   765
         End
         Begin VB.TextBox txtPrefMaxReply 
            Height          =   300
            Left            =   600
            TabIndex        =   17
            Text            =   "Replace Me"
            Top             =   720
            Width           =   495
         End
         Begin VB.CheckBox chkLimReply 
            Caption         =   "Limit reply length?"
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   360
            Width           =   2655
         End
         Begin VB.Label lblThreshold 
            Caption         =   "Threshold for start and end words (%)"
            Height          =   270
            Left            =   1500
            TabIndex        =   20
            Top             =   1200
            Width           =   2790
         End
         Begin VB.Label lblPrefMaxReply 
            Caption         =   "Preferred Maximum Reply Length"
            Height          =   270
            Left            =   1500
            TabIndex        =   19
            Top             =   780
            Width           =   2520
         End
      End
      Begin VB.Frame fraSample3 
         Caption         =   "Contextual"
         Height          =   1260
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   5175
         Begin VB.CheckBox chkExclude 
            Caption         =   "Ignore words in exclusion file?"
            Height          =   330
            Left            =   240
            TabIndex        =   21
            Top             =   675
            Width           =   2640
         End
         Begin VB.CheckBox chkContext 
            Caption         =   "Reply using contextual links?"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   330
            Width           =   2655
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
         Caption         =   "Contextualisation"
         Height          =   3225
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   5175
         Begin VB.TextBox txtMinConSize 
            Height          =   285
            Left            =   240
            MaxLength       =   4
            TabIndex        =   12
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Minimum word size for contextualisation"
            Height          =   255
            Left            =   840
            TabIndex        =   13
            Top             =   480
            Width           =   2895
         End
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   0
      Left            =   210
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample1 
         Caption         =   "Personal"
         Height          =   3225
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   5175
         Begin VB.TextBox txtName 
            Height          =   285
            Left            =   1320
            TabIndex        =   10
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label lblName 
            Caption         =   "Actor's Name:"
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   480
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
   Begin ComctlLib.TabStrip tbsProperties 
      Height          =   4245
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7488
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Personal"
            Key             =   "Personal"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Set Personal details for Actor"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Analysis"
            Key             =   "Analysis"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Set Options for Sentence Analysis"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Responses"
            Key             =   "Responses"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Set Options for Actor Responses"
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
Attribute VB_Name = "frmProperties"
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

Public Sub chkLimReply_Click()

    ' Enable or disable the associated reply length limit and threshold boxes depending on whether reply limiting is selected
    If chkLimReply = 1 Then
        txtPrefMaxReply.Enabled = True
        lblPrefMaxReply.Enabled = True
        cmbThreshold.Enabled = True
        lblThreshold.Enabled = True
    Else
        txtPrefMaxReply.Enabled = False
        lblPrefMaxReply.Enabled = False
        cmbThreshold.Enabled = False
        lblThreshold.Enabled = False
    End If
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()

    ' Invoke context help
    SendKeys "{F1}"
    
End Sub

Private Sub cmdOK_Click()

    On Error GoTo ErrorHandler

    Dim nMinConSize%

    ' Setup changes from dialog to data structures
    gobjCurrActor.strName = txtName
    
    ' Set contextual sizing
    ' Contextual word size should be some positive number
    If IsNumeric(txtMinConSize) Then
        gobjCurrActor.nMinConSize = CInt(txtMinConSize)
        txtMinConSize = gobjCurrActor.nMinConSize
    End If
    
    ' Set whether reply should be curtailed
    gobjCurrActor.bLimReplyLength = IIf(chkLimReply, True, False)
    
    ' Set preferred maximum reply length
    If IsNumeric(txtPrefMaxReply) Then
        gobjCurrActor.nPrefMaxReplyLength = CInt(txtPrefMaxReply)
        txtPrefMaxReply = gobjCurrActor.nPrefMaxReplyLength
    End If
    
    ' Set proportion of SOS or EOS links for a word to be considered a terminator
    If IsNumeric(cmbThreshold) Then
        gobjCurrActor.fTermProportionReq = CSng(cmbThreshold / 100)
        cmbThreshold = gobjCurrActor.fTermProportionReq * 100
    End If

    ' Transfer context stuff
    gobjCurrActor.bContextual = IIf(chkContext, True, False)
    gobjCurrActor.bExclude = IIf(chkExclude, True, False)
    
    ' Update the display
    frmMain.UpdateStatus
    
    Unload Me
    
Exit Sub
    
ErrorHandler:
    ' Catches an error which occurs if the input number is too big for an integer type
    ' Shouldn't be invoked for this because I've limited the box sizes
    
    MsgBox "Error in setting a property from the Actors Properties box"

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    'handle ctrl+tab to move to the next tab
    If Shift = vbCtrlMask And KeyCode = vbKeyTab Then
        i = tbsProperties.SelectedItem.Index
        If i = tbsProperties.Tabs.Count Then
            'last tab so we need to wrap to tab 1
            Set tbsProperties.SelectedItem = tbsProperties.Tabs(1)
        Else
            'increment the tab
            Set tbsProperties.SelectedItem = tbsProperties.Tabs(i + 1)
        End If
    End If
End Sub

Private Sub Form_Load()
    'center the form
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub

Private Sub tbsProperties_Click()
    
    Dim i As Integer
    'show and enable the selected tab's controls
    'and hide and disable all others
    For i = 0 To tbsProperties.Tabs.Count - 1
        If i = tbsProperties.SelectedItem.Index - 1 Then
            picOptions(i).Left = 210
            picOptions(i).Enabled = True
        Else
            picOptions(i).Left = -20000
            picOptions(i).Enabled = False
        End If
    Next
    
End Sub

Private Sub txtName_LostFocus()

    ' Sanity check - Make sure a blank name isn't entered
    If txtName = "" Then txtName = gobjCurrActor.strName
    
End Sub
