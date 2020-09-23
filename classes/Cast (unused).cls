VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsCast: Contains information pertaining to ALL actors.

Public mobjCurrActor As clsActor        ' Current operative actor
Public mcolExclusions As clsCollection  ' Contains words which we want to stop being treated as contextual

' LoadExclusions:
' Trust:
' Arguments: nFileNo - already open file
' Returns: NONE
Public Sub LoadExclusions(ByVal nFileNo As Integer)

End Sub

Private Sub Class_Initialize()

    ' Set up exclusions collection to accept only strings
    Set mcolExclusions = New clsCollection
    mcolExclusions.Initialise vbString
    
End Sub
