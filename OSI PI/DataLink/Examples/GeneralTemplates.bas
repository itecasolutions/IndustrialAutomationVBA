Attribute VB_Name = "GeneralTemplates"
'Written 02APR2021 by Nicholas Stom
'General Templates.
'These do not run in this module.
'These are templates to be copied and applied in other modules.


'Used for Cell Click to Open Form

Private Sub Worksheet_SelectionChange(ByVal Targe As Range)
    If ActiveCell.Value = "" Then
        On Error Resume Next
        With myFormName
            .StartUpPosition
            .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
            .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
            .Show
        End With
        On Error Resume Next
    End If
End Sub

 Sub WorksheetLoop()

         Dim WS_Count As Integer
         Dim I As Integer

         ' Set WS_Count equal to the number of worksheets in the active
         ' workbook.
         WS_Count = ActiveWorkbook.Worksheets.Count

         ' Begin the loop.
         For I = 1 To WS_Count

            ' Insert your code here.
            ' The following line shows how to reference a sheet within
            ' the loop by displaying the worksheet name in a dialog box.
            MsgBox ActiveWorkbook.Worksheets(I).name

         Next I

      End Sub
