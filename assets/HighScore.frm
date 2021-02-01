VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} HighScore 
   Caption         =   "High Scores"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8055
   OleObjectBlob   =   "HighScore.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "HighScore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NewScore As Integer

Private Sub ButtonClose_Click()
    
    Finish
    
End Sub

Private Sub UserForm_Activate()

    StartUp Score

End Sub

Private Sub UserForm_Initialize()
    
    StartUp
    
End Sub

Private Sub UserForm_Terminate()
    
    If NewScore > 0 Then
        
        Finish
        
    End If

End Sub

Private Function StartUp(Optional ByVal CurrentScore As Integer = 1000)
    
    Dim i As Integer, BeginningScores(), IntermediateScores(), ExpertScores() As Variant, wb As Workbook
    Diff = ""
    NewScore = 0
    
    Application.ScreenUpdating = False
    
    Set wb = Workbooks.Open("\\cllfw001\Rogers\Seniors\Dakota Blackwell\Projects\Misc\Scores.xlsm", ReadOnly:=False)
        
        ThisWorkbook.Sheets("BackEnd").Range("A2:F11") = wb.Sheets("Scores").Range("A2:F11").Value
        
    wb.Close False
    
    Application.ScreenUpdating = True
    
    BeginningScores = Application.Transpose(ThisWorkbook.Sheets("BackEnd").Range("A2:B11"))
    IntermediateScores = Application.Transpose(ThisWorkbook.Sheets("BackEnd").Range("C2:D11"))
    ExpertScores = Application.Transpose(ThisWorkbook.Sheets("BackEnd").Range("E2:F11"))
    
    If CurrentScore > 0 And CurrentScore < 1000 Then
        
        If Options.OptionBeginner.Value = True Then
            
            NewScore = AssignHighScore(BeginningScores, CurrentScore, "Beginner")
            
        ElseIf Options.OptionIntermediate.Value = True Then
            
            NewScore = AssignHighScore(IntermediateScores, CurrentScore, "Intermediate")
            
        ElseIf Options.OptionExpert.Value = True Then
            
            NewScore = AssignHighScore(ExpertScores, CurrentScore, "Expert")
            
        End If
        
    End If
    
    For i = 1 To 10
        
        Me.Controls("BeginnerScore" & i).Caption = CStr(BeginningScores(1, i))
        Me.Controls("BeginnerName" & i).Caption = CStr(BeginningScores(2, i))
        Me.Controls("IntermediateScore" & i).Caption = CStr(IntermediateScores(1, i))
        Me.Controls("IntermediateName" & i).Caption = CStr(IntermediateScores(2, i))
        Me.Controls("ExpertScore" & i).Caption = CStr(ExpertScores(1, i))
        Me.Controls("ExpertName" & i).Caption = CStr(ExpertScores(2, i))
        
    Next i
    
    If NewScore <> 0 Then
        
        Me.ButtonClose.Caption = "Submit"
        
    Else
        
        Me.ButtonClose.Caption = "Close"
        Me.BeginnerEnter.Visible = False
        Me.IntermediateEnter.Visible = False
        Me.ExpertEnter.Visible = False
        
    End If
    
End Function

Function AssignHighScore(ByRef ScoreArray() As Variant, ByVal CurrentScore As Integer, ByVal Diff As String) As Integer
    
    Dim i, Col As Integer
    
    If UBound(ScoreArray()) = 6 Then
        
        If Diff = "Intermediate" Then
            
            Col = 3
            
        ElseIf Diff = "Expert" Then
            
            Col = 5
            
        Else
            
            Col = 1
            
        End If
        
    Else
        
        Col = 1
        
    End If
    
    For i = 1 To 10
        
        If CurrentScore < ScoreArray(Col, i) Then
            
            Exit For
            
        End If
        
        If i = 10 Then
            
            AssignHighScore = 0
            Exit Function
            
        End If
        
    Next
    
    Dim j As Integer
    
    For j = 9 To i Step -1
        
        ScoreArray(Col, j + 1) = ScoreArray(Col, j)
        ScoreArray(Col + 1, j + 1) = ScoreArray(Col + 1, j)
        
    Next
    
    ScoreArray(Col, i) = CurrentScore
    ScoreArray(Col + 1, i) = ""
    
    Me.Controls(Diff & "Enter").Top = Me.Controls(Diff & "Name" & i).Top
    Me.Controls(Diff & "Enter").Visible = True
    AssignHighScore = i

End Function


Private Function Finish()
    
    If NewScore > 0 Then
        
        Application.ScreenUpdating = False
        
        Dim wb As Workbook, Row As Integer, ScoresArray() As Variant
        
        Set wb = Workbooks.Open("\\cllfw001\Rogers\Seniors\Dakota Blackwell\Projects\Misc\Scores.xlsm", ReadOnly:=False)
        
        ScoresArray = Application.Transpose(wb.Sheets(1).Range("A2:F11"))
                    
        If Me.BeginnerEnter.Visible = True Then
            
            Row = AssignHighScore(ScoresArray(), Score, "Beginner")
            ScoresArray(2, Row) = Me.BeginnerEnter.Text
            
        ElseIf Me.IntermediateEnter.Visible = True Then
            
            Row = AssignHighScore(ScoresArray(), Score, "Intermediate")
            ScoresArray(4, Row) = Me.IntermediateEnter.Text
            
        ElseIf Me.ExpertEnter.Visible = True Then
            
            Row = AssignHighScore(ScoresArray(), Score, "Expert")
            ScoresArray(6, Row) = Me.ExpertEnter.Text
            
        End If
        
        Application.Run "Scores.xlsm!UpdateScores", Application.Transpose(ScoresArray)
        
        wb.Close True
        
        Application.ScreenUpdating = True
        
        Score = 0
        StartUp
        
    Else
    
        HighScore.Hide
        
    End If
    
End Function

