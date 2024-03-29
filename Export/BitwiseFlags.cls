VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BitwiseFlags"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Stores bitwise flags between 1 and 30 inclusive."
'@IgnoreModule ProcedureNotUsed
'@ModuleDescription "Stores bitwise flags between 1 and 30 inclusive."
'@Folder "Helpers.Common"
Option Explicit

Private Const MAX_FLAG As Long = 30
Private Const ERR_NUMBER As Long = 1
Private Const ERR_MSG_OUTOFBOUNDS As String = "BitwiseFlags class only supports flags between 1 and 30 inclusive"

Private Type TState
    State As Long
End Type
Private This As TState

Public Property Get State() As Long
    State = This.State
End Property

Public Function Has(ByVal Flag As Long) As Boolean
    GuardFlagOutOfBounds Flag
    Has = (This.State And (2 ^ Flag)) = (2 ^ Flag)
End Function

Public Sub Add(ByVal Flag As Long)
    GuardFlagOutOfBounds Flag
    If Has(Flag) Then Exit Sub
    This.State = This.State + (2 ^ Flag)
End Sub

Public Sub Remove(ByVal Flag As Long)
    GuardFlagOutOfBounds Flag
    If Not Has(Flag) Then Exit Sub
    This.State = This.State - (2 ^ Flag)
End Sub

Public Function Equals(ByVal RHS As BitwiseFlags) As Boolean
    Equals = (This.State = RHS.State)
End Function

Public Sub DebugPrint()
    Debug.Print "BitwiseFlags State = "; This.State
    
    Dim MaxFlag As Long
    MaxFlag = Sqr(This.State)
    If MaxFlag > MAX_FLAG Then MaxFlag = MAX_FLAG
    
    Dim i As Long
    For i = 1 To MaxFlag
        If Has(i) Then
            Debug.Print " "; Format$(i, "00"); " = SET"
        End If
    Next i
End Sub

Private Sub GuardFlagOutOfBounds(ByVal Flag As Long)
    If Flag <= 0 Or Flag > MAX_FLAG Then
        Err.Raise vbObjectError + ERR_NUMBER, "BitwiseFlags", ERR_MSG_OUTOFBOUNDS
    End If
End Sub
