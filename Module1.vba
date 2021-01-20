Option Explicit

Sub Run_Quiz()
    'Set Title and Description to Values in Sheet.
    Range("A1").Select
    TitlePage.Title.Caption = ActiveCell.Value
    TitlePage.Description.Caption = ActiveCell.Offset(1).Value
    
    'Determine the total number of questions in the sheet and adjust labels accordingly.
    Range("A11").Select
    Dim MaxQuestions As Integer
    Dim i As Integer
    MaxQuestions = 0
    While ActiveCell.Offset(i).Value <> ""
        MaxQuestions = MaxQuestions + 1
        i = i + 1
    Wend
    TitlePage.Max = MaxQuestions & ")"
    
    'Shuffle the questions in the sheet.
    Call ShuffleQuestions
    
    'Display the Title Page of the quiz.
    TitlePage.Show
End Sub

Private Sub ShuffleQuestions()
    'Set values in column J (default white text) to random values between 0 and 1000.
    Range("J11").Select
    While ActiveCell.Offset(, -9).Value <> ""
        ActiveCell.Value = Application.WorksheetFunction.RandBetween(0, 1000)
        ActiveCell.Offset(1).Select
    Wend
    ActiveCell.Offset(-1).Select
    
    'Sort question table based on column J
    Dim BottomRight As String
    BottomRight = ActiveCell.Address
    Range("A11", BottomRight).Sort Key1:=Range("J11"), Order1:=xlAscending, Header:=xlNo
End Sub

Sub Clear_Results()
    'Empty cells containing score
    Range("B4").Value = ""
    Range("C4").Value = ""
    
    'Remove fill colors and empty column I (containing points earned)
    Range("A11").Select
    Dim row As Integer
    row = 0
    While ActiveCell.Offset(row).Value <> ""
        Range(ActiveCell.Offset(row), ActiveCell.Offset(row, 8)).Interior.Color = xlNone
        ActiveCell.Offset(row, 8).Value = ""
        row = row + 1
    Wend
End Sub