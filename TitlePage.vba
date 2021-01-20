Option Explicit

Private Sub Start_Click()
    'Determine the user inputted number of questions
    Dim UserInput As String
    Dim DesiredQuestions As Integer
    UserInput = TitlePage.NumberQuestions.Text
    If IsNumeric(UserInput) Then
        DesiredQuestions = CInt(UserInput)
    Else
        DesiredQuestions = 0
    End If
    
    'Determine the maximum number of questions
    Dim MaxLabel As String
    Dim MaxLabelSubstr As String
    Dim ParPos As Integer
    Dim MaxNum As Integer
    MaxLabel = TitlePage.Max.Caption
    ParPos = InStr(1, MaxLabel, ")")
    MaxLabelSubstr = Left(MaxLabel, ParPos - 1)
    MaxNum = CInt(MaxLabelSubstr)
    
    'Print error message if invalid value is inputted.
    'Otherwise, set appropriate values and open the next form.
    If TitlePage.NumberQuestions.Text = "" Then
        MsgBox ("Please input desired number of questions in the text box.")
    ElseIf DesiredQuestions <= 0 Or DesiredQuestions > MaxNum Then
        MsgBox ("Please input a valid number of questions.")
    Else
        Range("B5").Value = 1
        Range("C5").Value = DesiredQuestions
        Range("D5").Value = 0
        Range("E5").Value = 0
        Question.Caption = "Question " & Range("B5").Value & " of " & Range("C5").Value
        Range("A11").Select
        
        Call SetText
        
        Unload Me
        Question.Show
    End If
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
