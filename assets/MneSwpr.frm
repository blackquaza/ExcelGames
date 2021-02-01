VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Game 
   Caption         =   "Minesweeper"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9210
   OleObjectBlob   =   "MneSwpr.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Game"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ButtonDownX, ButtonDownY As Integer

Private Sub ButtonRestart_Click()
    
    If Not Won And Not Lost Then Restart = True
    Call LoadGame
    
End Sub

'Private Sub UserForm_Activate()
'
'    Call LoadGame
'
'End Sub

Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    
    DoubleClick ButtonDownY, ButtonDownX

End Sub

Private Sub UserForm_Deactivate()
    
    If Not Won Then
        
        Lose
        
    End If
    
End Sub

Private Sub UserForm_Initialize()
    
    Call LoadGame
    
End Sub

Function LoadGame()
    
    First = True
    Won = False
    Lost = False
    Game.TimeCount.Caption = 0
    
    Dim i, j As Integer
    
    ReDim GameField(1 To MaxY, 1 To MaxX)
    ReDim ButtonArray(1 To MaxY, 1 To MaxX)
    ReDim LabelArray(1 To MaxY, 1 To MaxX)
    
    Remaining = MaxY * MaxX
    Flags = Mines
    Game.MineCount.Caption = Flags
    
    For i = 1 To 24
        
        For j = 1 To 30
            
            If i <= MaxY And j <= MaxX Then
                
                Game.Controls("P_" & i & "_" & j).Visible = True
                Game.Controls("P_" & i & "_" & j).Picture = Game.ButtonUp.Picture
                
            Else
                
                Game.Controls("P_" & i & "_" & j).Visible = False
                
            End If
            
        Next
        
    Next
    
    Game.Width = MaxX * 15 + 15
    Game.Height = MaxY * 15 + 90
    Game.TextTimer.Left = Game.Width - 66
    Game.ButtonRestart.Left = Game.Width / 2 - 26.3
    
End Function

Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    ButtonDownX = Int((X + 10) / 15)
    ButtonDownY = Int((Y - 50) / 15)
    
    If ButtonDownX >= 1 And ButtonDownX <= MaxX And ButtonDownY >= 1 And ButtonDownY <= MaxY And _
    Button = 1 And Not Lost And Not Won Then
        
        If Game.Controls("P_" & ButtonDownY & "_" & ButtonDownX).Picture = Game.ButtonUp.Picture Or _
        Game.Controls("P_" & ButtonDownY & "_" & ButtonDownX).Picture = Game.ButtonMaybe.Picture Then
        
            Game.Controls("P_" & ButtonDownY & "_" & ButtonDownX).Picture = Game.ButtonDown.Picture
            
        End If
        
    End If
    
End Sub

Private Sub UserForm_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    If ButtonDownX >= 1 And ButtonDownX <= MaxX And ButtonDownY >= 1 And ButtonDownY <= MaxY And _
    Not Lost And Not Won Then
        
        If Int((X + 10) / 15) = ButtonDownX And Int((Y - 50) / 15) = ButtonDownY Then
            
            If Button = 1 Then
                    
                ButtonPress ButtonDownY, ButtonDownX
                     
            ElseIf Button = 2 Then
                
                RightClick ButtonDownY, ButtonDownX
                
            End If
            
        Else
            
            If Game.Controls("P_" & ButtonDownY & "_" & ButtonDownX).Picture = Game.ButtonDown.Picture Then
            
                Game.Controls("P_" & ButtonDownY & "_" & ButtonDownX).Picture = Game.ButtonUp.Picture
                
            End If
            
        End If
        
    End If
    
End Sub

Private Sub UserForm_Terminate()
    
    If Not Won Then
        
        Lose
        
    End If

End Sub
