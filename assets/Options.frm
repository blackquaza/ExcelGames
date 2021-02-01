VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Options 
   Caption         =   "Minesweeper"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5100
   OleObjectBlob   =   "Options.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ButtonHighScore_Click()
    
    'ThisWorkbook.Save
    HighScore.Show
    
End Sub

Private Sub ButtonStart_Click()
    
    CustomCheck
    'ThisWorkbook.Save
    Game.Show
    
End Sub

Private Sub InputLength_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    
    CustomCheck
    
End Sub

Private Sub InputMines_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    
    CustomCheck
    
End Sub

Private Sub InputHeight_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    
    CustomCheck
    
End Sub

Function CustomCheck()
    
    If Not IsNumeric(Options.InputHeight.Text) Then
        
        Options.InputHeight.Text = 9
        
    ElseIf Options.InputHeight.Text < 8 Then
        
        Options.InputHeight.Text = 8
        
    ElseIf Options.InputHeight.Text > 24 Then
        
        Options.InputHeight.Text = 24
        
    ElseIf Options.InputHeight.Text <> Int(Options.InputHeight.Text) Then
        
        Options.InputHeight.Text = Int(Options.InputHeight.Text)
        
    End If
    
    MaxY = Int(Options.InputHeight.Text)
    
    '---
    
    If Not IsNumeric(Options.InputLength.Text) Then
        
        Options.InputLength.Text = 9
        
    ElseIf Options.InputLength.Text < 8 Then
        
        Options.InputLength.Text = 8
        
    ElseIf Options.InputLength.Text > 30 Then
        
        Options.InputLength.Text = 30
        
    ElseIf Options.InputLength.Text <> Int(Options.InputLength.Text) Then
        
        Options.InputLength.Text = Int(Options.InputLength.Text)
        
    End If
    
    MaxX = Int(Options.InputLength.Text)
    
    '---
    
    If Not IsNumeric(Options.InputMines.Text) Then
        
        Options.InputMines.Text = RoundUp(MaxY * MaxX / 20)
        
    ElseIf Int(Options.InputMines.Text) < RoundUp(MaxY * MaxX / 10) Then
        
        Options.InputMines.Text = RoundUp(MaxY * MaxX / 20)
        
    ElseIf Int(Options.InputMines.Text) > ((MaxX - 1) * (MaxY - 1)) Then
        
        Options.InputMines.Text = (MaxX - 1) * (MaxY - 1)
        
    ElseIf Options.InputMines.Text <> Int(Options.InputMines.Text) Then
        
        Options.InputMines.Text = Int(Options.InputMines.Text)
        
    End If
    
    Mines = Int(Options.InputMines.Text)
    
End Function

Private Sub OptionBeginner_Click()
    
    Options.InputLength.Text = 9
    Options.InputHeight.Text = 9
    Options.InputMines.Text = 10
    
    Options.InputLength.Enabled = False
    Options.InputHeight.Enabled = False
    Options.InputMines.Enabled = False
    
End Sub

Private Sub OptionCustom_Click()
    
    Options.InputLength.Enabled = True
    Options.InputHeight.Enabled = True
    Options.InputMines.Enabled = True
    
End Sub

Private Sub OptionExpert_Click()
    
    Options.InputLength.Text = 30
    Options.InputHeight.Text = 16
    Options.InputMines.Text = 99
    
    Options.InputLength.Enabled = False
    Options.InputHeight.Enabled = False
    Options.InputMines.Enabled = False
    
End Sub

Private Sub OptionIntermediate_Click()
    
    Options.InputLength.Text = 16
    Options.InputHeight.Text = 16
    Options.InputMines.Text = 40
    
    Options.InputLength.Enabled = False
    Options.InputHeight.Enabled = False
    Options.InputMines.Enabled = False
    
End Sub

Private Sub Tutorial_Click()
    
    MsgBox "Wow. You fail."
    
End Sub

Private Sub UserForm_Initialize()
    
    Options.InputLength.Text = 9
    Options.InputHeight.Text = 9
    Options.InputMines.Text = 10
    
    Options.InputLength.Enabled = False
    Options.InputHeight.Enabled = False
    Options.InputMines.Enabled = False
    
    Score = 0
    
End Sub

Private Sub UserForm_Terminate()
    
    ActiveWindow.WindowState = State
    
End Sub
