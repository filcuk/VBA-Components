Attribute VB_Name = "basConvertFormulaReference"
Option Explicit

Sub MakeAbsoluteorRelativeFast()
'Original written by OzGrid Business Applications
'www.ozgrid.com
'Edited 2017-03-02

Dim RdoRange As Range
Dim i As Integer
Dim Reply As String

    'Ask whether Relative or Absolute
    Reply = InputBox("Change formulas to?" & Chr(13) & Chr(13) _
    & "Relative row/Absolute column = 1" & Chr(13) _
    & "Absolute row/Relative column = 2" & Chr(13) _
    & "Absolute all = 3" & Chr(13) _
    & "Relative all = 4", "OzGrid Business Applications")

    'They cancelled
    If Reply = "" Then Exit Sub
    
    On Error Resume Next
    
    'Set Range variable to formula cells only
    Set RdoRange = Selection.SpecialCells(Type:=xlFormulas)
    
    'determine the change type
    Select Case Reply
    
        Case 1 'Relative row/Absolute column
        
            For i = 1 To RdoRange.Areas.Count
                RdoRange.Areas(i).Formula = _
                Application.ConvertFormula _
                (Formula:=RdoRange.Areas(i).Formula, _
                FromReferenceStyle:=xlA1, _
                ToReferenceStyle:=xlA1, ToAbsolute:=xlRelRowAbsColumn)
            Next i
            
        Case 2 'Absolute row/Relative column
        
            For i = 1 To RdoRange.Areas.Count
                RdoRange.Areas(i).Formula = _
                Application.ConvertFormula _
                (Formula:=RdoRange.Areas(i).Formula, _
                FromReferenceStyle:=xlA1, _
                ToReferenceStyle:=xlA1, ToAbsolute:=xlAbsRowRelColumn)
            Next i
            
        Case 3 'Absolute all
        
            For i = 1 To RdoRange.Areas.Count
                RdoRange.Areas(i).Formula = _
                Application.ConvertFormula _
                (Formula:=RdoRange.Areas(i).Formula, _
                FromReferenceStyle:=xlA1, _
                ToReferenceStyle:=xlA1, ToAbsolute:=xlAbsolute)
            Next i
            
        Case 4 'Relative all
            For i = 1 To RdoRange.Areas.Count
                RdoRange.Areas(i).Formula = _
                Application.ConvertFormula _
                (Formula:=RdoRange.Areas(i).Formula, _
                 FromReferenceStyle:=xlA1, _
                 ToReferenceStyle:=xlA1, ToAbsolute:=xlRelative)
            Next i
            
        Case Else 'Typo
            MsgBox "Change type not recognised!", vbCritical, _
            "OzGrid Business Applications"
    End Select
    
    'Clear memory
    Set RdoRange = Nothing
End Sub
