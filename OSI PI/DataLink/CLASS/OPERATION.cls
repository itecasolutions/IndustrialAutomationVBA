VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OPERATION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Written 22APR2021 by Nicholas Stom
'Class is used to Get Unit Procedure from OSI PI from DeltaV Run on Unit
'


Option Explicit
' These are properties
    Public startTime As Date
    Public endTime As Date
    Public piTag As Variant
    Public unitName As String
    Public dataMatrix As Variant
    Public timedPhaseId As Variant
    Public timedBatchId As Variant
    
Public Sub GetOperationStyle1()
   'Data Precueded
   '01-JAN-21 12:00:00  OPERATION:1-1
   '01-JAN-21 12:00:00  Inactive


       'SET PAGE CHARACTERISTICS
           Application.Calculation = xlCalculationManual
           Application.ScreenUpdating = False
       'DECLARE VARIABLE
           Dim substrings() As String
           Dim unformatted_matrix, my_matrix, batch_mtx, task_mtx As Variant
           Dim i, j As Integer
           Dim LString As String
           Dim lArray() As String
           Dim procedural_modal() As String
           Dim task_tag As String
           Dim sizeFlag As Integer
           Dim firstPass, toggleStart As Boolean
           startTime = CDate(startTime)
           endTime = CDate(endTime)
   
           'DEFINE VARIABLES
               ReDim my_matrix(20, 0)
               sizeFlag = -1
               firstPass = False
               toggleStart = True
               unformatted_matrix = Application.Run("PICompDat", piTag, startTime, endTime, 9, "", "inside")
                       'Move to my_matrix
                           On Error Resume Next
                           If unformatted_matrix(1) <> "No more values:" Then
                           On Error Resume Next
                               For i = 1 To UBound(unformatted_matrix, 1)
                                   If unformatted_matrix(i, 1) <> " " And unformatted_matrix(1) <> "No more values:" Then
                                       If unformatted_matrix(i, 2) <> "Inactive" And unformatted_matrix(i, 1) <> "Resize to show all values" Then
                                           sizeFlag = sizeFlag + 1
                                           ReDim Preserve my_matrix(20, sizeFlag)
                                           my_matrix(0, sizeFlag) = unitName
                                           my_matrix(1, sizeFlag) = unformatted_matrix(i, 2)
                                           my_matrix(2, sizeFlag) = Format(unformatted_matrix(i, 1), "mm/dd/yyyy hh:mm:ss")
                                            If unformatted_matrix(i + 1, 2) = "Inactive" And (i + 1) <= UBound(unformatted_matrix, 1) And unformatted_matrix(i + 1, 1) <> "Resize to show all values" Then
                                                my_matrix(3, sizeFlag) = Format(unformatted_matrix(i + 1, 1), "mm/dd/yyyy hh:mm:ss")
                                            ElseIf unformatted_matrix(i + 2, 2) = "Inactive" And (i + 2) <= UBound(unformatted_matrix, 1) And unformatted_matrix(i + 2, 1) <> "Resize to show all values" Then
                                                my_matrix(3, sizeFlag) = Format(unformatted_matrix(i + 2, 1), "mm/dd/yyyy hh:mm:ss")
                                            ElseIf unformatted_matrix(i + 3, 2) = "Inactive" And (i + 3) <= UBound(unformatted_matrix, 1) And unformatted_matrix(i + 3, 1) <> "Resize to show all values" Then
                                                my_matrix(3, sizeFlag) = Format(unformatted_matrix(i + 3, 1), "mm/dd/yyyy hh:mm:ss")
                                            ElseIf unformatted_matrix(i + 4, 2) = "Inactive" And (i + 4) <= UBound(unformatted_matrix, 1) And unformatted_matrix(i + 4, 1) <> "Resize to show all values" Then
                                                my_matrix(3, sizeFlag) = Format(unformatted_matrix(i + 4, 1), "mm/dd/yyyy hh:mm:ss")
                                            End If
                                       End If

                                   End If
                               Next i
                           Else
                           End If

       dataMatrix = my_matrix
End Sub

