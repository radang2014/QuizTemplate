Option Explicit

Private Sub Option1_Click()
    Call CheckCorrect(1)
End Sub

Private Sub Option2_Click()
    Call CheckCorrect(2)
End Sub

Private Sub Option3_Click()
    Call CheckCorrect(3)
End Sub

Private Sub Option4_Click()
    Call CheckCorrect(4)
End Sub

Private Sub Option5_Click()
    Call CheckCorrect(5)
End Sub

Private Sub CheckCorrect(Selected As Integer)
    'Adds one to the value storing the question number.
    Range("B5").Value = Range("B5").Value + 1
    
    'Check if selected answer is correct and calls appropriate subroutine.
    If Selected = ActiveCell.Offset(, 6).Value Then
        Call Correct
    Else
        Call Incorrect
    End If
    
    'Selects the next question in the sheet.
    ActiveCell.Offset(1).Select
    
    'Change the text in the user form to that representing the next question.
    Question.Caption = "Question " & Range("B5").Value & " of " & Range("C5").Value
    Call SetText
    
    'If the user reached the last question, determine the score and display the score page.
    If Range("B5").Value > Range("C5").Value Then
        Dim Score As String
        Dim Percentage As Double
        Score = Range("D5").Value & " / " & Range("E5").Value
        Percentage = Range("D5").Value / Range("E5").Value
        
        ScorePage.Score.Caption = Score
        ScorePage.Percentage.Caption = Round(Percentage * 100, 2) & "%"
        
        Range("B4").Value = "'" & Score
        Range("C4").Value = Percentage
        Unload Me
        ScorePage.Show
    End If
End Sub

Private Sub Correct()
    'Highlight the row green
    Range(ActiveCell.Address, ActiveCell.Offset(, 8).Address).Interior.Color = RGB(0, 255, 0)
    
    'Determines the point value of the question, which is the number of points earned for answering correctly.
    Dim PointValue As Double
    PointValue = ActiveCell.Offset(, 7).Value
    ActiveCell.Offset(, 8).Value = PointValue
    
    'Adds the point value to both the numerator and denominator of the score.
    Range("D5").Value = Range("D5").Value + PointValue
    Range("E5").Value = Range("E5").Value + PointValue
End Sub

Private Sub Incorrect()
    'Highlight the row red
    Range(ActiveCell.Address, ActiveCell.Offset(, 8).Address).Interior.Color = RGB(255, 0, 0)
    
    'Determines the point value of the question; user earns 0 points for answering incorrectly.
    Dim PointValue As Double
    PointValue = ActiveCell.Offset(, 7).Value
    ActiveCell.Offset(, 8).Value = 0
    
    'Adds the point value to the denominator of the score
    Range("E5").Value = Range("E5").Value + PointValue
End Sub

Private Sub SetText()
    'Set captions to values in the sheet
    Question.QuestionText.Caption = ActiveCell.Value
    Question.Option1.Caption = "A. " & ActiveCell.Offset(, 1).Value
    Question.Option2.Caption = "B. " & ActiveCell.Offset(, 2).Value
    Question.Option3.Caption = "C. " & ActiveCell.Offset(, 3).Value
    Question.Option4.Caption = "D. " & ActiveCell.Offset(, 4).Value
    Question.Option5.Caption = "E. " & ActiveCell.Offset(, 5).Value
    
    'If there is no text for any option, do not show the option.
    If ActiveCell.Offset(, 1).Value = "" Then
        Question.Option1.Visible = False
    Else
        Question.Option1.Visible = True
    End If
    
    If ActiveCell.Offset(, 2).Value = "" Then
        Question.Option2.Visible = False
    Else
        Question.Option2.Visible = True
    End If
    
    If ActiveCell.Offset(, 3).Value = "" Then
        Question.Option3.Visible = False
    Else
        Question.Option3.Visible = True
    End If
    
    If ActiveCell.Offset(, 4).Value = "" Then
        Question.Option4.Visible = False
    Else
        Question.Option4.Visible = True
    End If
    
    If ActiveCell.Offset(, 5).Value = "" Then
        Question.Option5.Visible = False
    Else
        Question.Option5.Visible = True
    End If
End Sub