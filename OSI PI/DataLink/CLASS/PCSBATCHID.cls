VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PCSBATCHID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Written 24AUG2020 by Nicholas Stom
'Class is used to Get PCS Batch IDs Run on Unit
'   The PiTag is a public variant which means it can be an array of multiple units

Option Explicit
' These are properties
    Public startTime As Date
    Public endTime As Date
    Public piTag As Variant
    Public dataMatrix As Variant
    Public batchId As String
    Public unitName As Variant
' These are methods
    Public Sub Get_PcsBatchId()
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
                If IsArray(piTag) = True Then
                    For j = 0 To UBound(piTag)
                        unformatted_matrix = Application.Run("PICompDat", piTag(j), startTime, endTime, 9, "", "inside")
                        'Move to my_matrix
                            On Error Resume Next
                            If unformatted_matrix(1) <> "No more values:" Then
                            On Error Resume Next
                                For i = 1 To UBound(unformatted_matrix, 1)
                                    If unformatted_matrix(i, 1) <> " " And unformatted_matrix(1) <> "No more values:" Then
                                        If toggleStart = True And InStr(unformatted_matrix(i, 2), "UNITBATCH") > 0 And InStr(unformatted_matrix(i, 2), "START") > 0 Then
                                            sizeFlag = sizeFlag + 1
                                            ReDim Preserve my_matrix(20, sizeFlag)
                                            lArray = Split(piTag(j), "_")
                                            my_matrix(0, sizeFlag) = lArray(2)
                                            lArray = Split(unformatted_matrix(i, 2), "|")
                                            my_matrix(1, sizeFlag) = lArray(0)
                                            my_matrix(2, sizeFlag) = Format(unformatted_matrix(i, 1), "mm/dd/yyyy hh:mm:ss")
                                            firstPass = True
                                            toggleStart = False
                                        End If
                                        lArray = Split(unformatted_matrix(i, 2), "|")
                                        If toggleStart = False And firstPass = True And InStr(unformatted_matrix(i, 2), "UNITBATCH") > 0 And InStr(unformatted_matrix(i, 2), "END") > 0 Then
                                            my_matrix(3, sizeFlag) = Format(unformatted_matrix(i, 1), "mm/dd/yyyy hh:mm")
                                            toggleStart = True
                                        End If
                                    End If
                                Next i
                            Else
                            End If
                    Next j
                Else
                    unformatted_matrix = Application.Run("PICompDat", piTag, startTime, endTime, 9, "", "inside")
                        'Move to my_matrix
                            On Error Resume Next
                            If unformatted_matrix(1) <> "No more values:" Then
                            On Error Resume Next
                                For i = 1 To UBound(unformatted_matrix, 1)
                                    If unformatted_matrix(i, 1) <> " " And unformatted_matrix(1) <> "No more values:" Then
                                        If InStr(unformatted_matrix(i, 2), "UNITBATCH") > 0 And InStr(unformatted_matrix(i, 2), "START") > 0 Then
                                            sizeFlag = sizeFlag + 1
                                            ReDim Preserve my_matrix(20, sizeFlag)
                                            lArray = Split(piTag, "_")
                                            my_matrix(0, sizeFlag) = lArray(2)
                                            lArray = Split(unformatted_matrix(i, 2), "|")
                                            my_matrix(1, sizeFlag) = lArray(0)
                                            my_matrix(2, sizeFlag) = Format(unformatted_matrix(i, 1), "mm/dd/yyyy hh:mm")
                                            firstPass = True
                                            toggleStart = False
                                        End If
                                        If firstPass = True And InStr(unformatted_matrix(i, 2), "UNITBATCH") > 0 And InStr(unformatted_matrix(i, 2), "END") > 0 Then
                                            my_matrix(3, sizeFlag) = Format(unformatted_matrix(i, 1), "mm/dd/yyyy hh:mm")
                                            toggleStart = True
                                            
                                        End If
                                    End If
                                Next i
                            Else
                            End If
                End If

        dataMatrix = my_matrix
    End Sub
    
    Public Sub GetBatchIdAtTime()
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
            unformatted_matrix = Application.Run("PIArcVal", piTag, startTime, 0, "", "auto")

            'Move to my_matrix
                On Error Resume Next
                If unformatted_matrix(1) <> "No more values:" Then
                On Error Resume Next
                    lArray = Split(unformatted_matrix(1), "|")
                    batchId = lArray(0)
                End If

    End Sub

    Public Sub GetBatchIdNoSplit()
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
                                        If unformatted_matrix(i, 2) <> " " Then
                                            sizeFlag = sizeFlag + 1
                                            ReDim Preserve my_matrix(20, sizeFlag)
                                            my_matrix(0, sizeFlag) = unitName
                                            my_matrix(1, sizeFlag) = unformatted_matrix(i, 2)
                                            my_matrix(2, sizeFlag) = Format(unformatted_matrix(i, 1), "mm/dd/yyyy hh:mm:ss")
                                            If (i + 1) < UBound(unformatted_matrix, 1) Then
                                                my_matrix(3, sizeFlag) = Format(unformatted_matrix(i + 1, 1), "mm/dd/yyyy hh:mm:ss")
                                            End If
                                        End If

                                    End If
                                Next i
                            Else
                            End If

        dataMatrix = my_matrix
    End Sub


Public Sub GetBatchIdStyle3()
   'Data Precueded
   '01-JAN-21 12:00:00 Batch ID
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
                                           my_matrix(2, sizeFlag) = CDate(Format(unformatted_matrix(i, 1), "mm/dd/yyyy hh:mm:ss"))
                                           If (i + 1) < UBound(unformatted_matrix, 1) And unformatted_matrix(i + 1, 1) <> "Resize to show all values" Then
                                               my_matrix(3, sizeFlag) = CDate(Format(unformatted_matrix(i + 1, 1), "mm/dd/yyyy hh:mm:ss"))
                                           ElseIf (i + 2) < UBound(unformatted_matrix, 1) And unformatted_matrix(i + 2, 1) <> "Resize to show all values" Then
                                                 my_matrix(3, sizeFlag) = CDate(Format(unformatted_matrix(i + 2, 1), "mm/dd/yyyy hh:mm:ss"))
                                           End If
                                       End If

                                   End If
                               Next i
                           Else
                           End If

       dataMatrix = my_matrix
End Sub


Public Sub CombinedBatchId()
   'Written 24MAY2021 by Nicholas Stom
   'Combined Batch IDs if Multiple Batch IDs
   'Typcially do to multiple unit procedures in a procedure


       'SET PAGE CHARACTERISTICS
           Application.Calculation = xlCalculationManual
           Application.ScreenUpdating = False
       'DECLARE VARIABLE
           Dim my_matrix As Variant
           Dim i, j, sizeFlag As Integer
           Dim currentId As String
  
   
           'DEFINE VARIABLES
                ReDim my_matrix(20, 0)
                sizeFlag = -1
                currentId = ""
                
            For j = 0 To UBound(dataMatrix, 2)
                If currentId <> dataMatrix(1, j) Then
                    sizeFlag = sizeFlag + 1
                    ReDim Preserve my_matrix(20, sizeFlag)
                    my_matrix(0, sizeFlag) = dataMatrix(0, j)
                    my_matrix(1, sizeFlag) = dataMatrix(1, j)
                    my_matrix(2, sizeFlag) = dataMatrix(2, j)
                    my_matrix(3, sizeFlag) = dataMatrix(3, j)
                    currentId = dataMatrix(1, j)
                ElseIf currentId = dataMatrix(1, j) Then
                    my_matrix(3, sizeFlag) = dataMatrix(3, j)
                End If
            Next j
       dataMatrix = my_matrix
End Sub

