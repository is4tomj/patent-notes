Renumber Claims in MS Word
===========================

## Install

In the Developer tab, open Visual Basic. Create a new macro with the following code:

```vb
Sub ChangeNumbers()
'
' IncClaims Macro
'

    ' get selected text
    Dim Sel As Selection
    Set Sel = Application.Selection
    If Len(Sel) < 2 Then
        MsgBox "No text block selected"
        Exit Sub
    End If
    
    
    ' the starting number
    Dim StartingNumberString As String
    Dim StartingNumber As Integer
    StartingNumberString = InputBox( _
        "Start at number", _
        "IncClaims", _
        "")
    StartingNumber = ConvertToInteger(StartingNumberString)
    

    ' the starting number
    Dim EndingNumberString As String
    Dim EndingNumber As Integer
    EndingNumberString = InputBox( _
        "Ending at number", _
        "IncClaims", _
        "")
    EndingNumber = ConvertToInteger(EndingNumberString)


    ' the change number
    Dim ChangeNumberString As String
    Dim ChangeNumber As Integer
    ChangeNumberString = InputBox( _
        "Change by", _
        "ChangeClaims", _
        "")
    ChangeNumber = ConvertToInteger(ChangeNumberString)
    

    Dim J As Integer
    Dim sFindText As String
    Dim sReplaceText As String

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    
    If ChangeNumber > 0 Then
        For J = EndingNumber To StartingNumber Step -1
            sFindText = J
            sReplaceText = J + ChangeNumber
            Selection.Find.Text = sFindText
            Selection.Find.Replacement.Text = sReplaceText
            Selection.Find.Execute Replace:=wdReplaceAll
        Next J
    End If
    If ChangeNumber < 0 Then
        For J = StartingNumber To EndingNumber Step 1
            sFindText = J
            sReplaceText = J + ChangeNumber
            Selection.Find.Text = sFindText
            Selection.Find.Replacement.Text = sReplaceText
            Selection.Find.Execute Replace:=wdReplaceAll
        Next J
    End If



End Sub

Function ConvertToInteger(v1 As Variant) As Integer
    On Error GoTo 100:
         ConvertToInteger = CInt(v1)
         Exit Function
100:
         MsgBox "Failed to convert """ & v1 & """ to an integer.", , "Aborting - Failed Conversion"
     End
End Function
```

## Usage

1.  Select the claims you want to renumber, and run the new macro – ChangeNumbers.
2.  Input the lowest claim number to modify.
3.  Input the highest claim number to modify.
4.  Input the amount (as a positive or negative integer) that you want to change each number within the selected text that is also within the inputed range.

