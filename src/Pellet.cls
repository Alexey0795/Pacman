VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pellet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PacmanGame.Model.Maps"
Option Explicit

'@MemberAttribute VB_VarHelpID, -1
Public WithEvents HostCell As Tile
Attribute HostCell.VB_VarHelpID = -1
Public parent As Collection

Private Sub HostCell_DispatchEncounter()
    '// remove the pellet
    '//Decrement pellets from the game and apply score
    If Not parent Is Nothing Then
        parent.Remove (Me.ToString)
        
    End If
End Sub

Public Function ToString() As String
    ToString = Drawing.TextDot
End Function
