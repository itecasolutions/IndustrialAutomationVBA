VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PHASE"
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
    Public unitParseChar As String
    Public unitPosition As Integer
    Public dataMatrix As Variant
    Public timedPhaseId As Variant
    Public timedBatchId As Variant
    


    Public Sub Get_PcsPhase()
        'SET PAGE CHARACTERISTICS
            Application.Calculation = xlCalculationManual
            Application.ScreenUpdating = False
        'DECLARE VARIABLE
            Dim substrings() As String
            Dim unformatted_matrix, my_matrix, phase1_mtx, phase2_mtx, batch_mtx, task_mtx As Variant
            Dim i, j As Integer
            Dim LString As String
            Dim lArray() As String
            Dim procedural_modal() As String
            Dim task_tag As String
            Dim sizeFlag, sizeFlag1 As Integer
            Dim firstPass, toggleStart As Boolean
            startTime = CDate(startTime)
            endTime = CDate(endTime)
    
            'DEFINE VARIABLES
                ReDim my_matrix(20, 0)
                ReDim phase1_mtx(3, 0)
                ReDim phase2_mtx(3, 0)
                sizeFlag = -1
                sizeFlag1 = -1
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
                                        If InStr(unformatted_matrix(i, 2), "PHASE|") > 0 And InStr(unformatted_matrix(i, 2), "START") > 0 Then
                                           'Resize Array
                                                sizeFlag = sizeFlag + 1
                                                ReDim Preserve phase1_mtx(3, sizeFlag)
                                            'Set Unit
                                                lArray = Split(piTag(j), unitParseChar)
                                                phase1_mtx(0, sizeFlag) = lArray(unitPosition)
                                            'Batch ID
                                                lArray = Split(unformatted_matrix(i, 2), "|")
                                                phase1_mtx(1, sizeFlag) = lArray(0)
                                            'PHASE
                                                lArray = Split(unformatted_matrix(i, 2), "|")
                                                lArray = Split(lArray(1), "\")
                                                lArray = Split(lArray(3), ":")
                                                phase1_mtx(2, sizeFlag) = lArray(0)
                                            'Time
                                                phase1_mtx(3, sizeFlag) = Format(unformatted_matrix(i, 1), "mm/dd/yyyy hh:mm:ss")
                                        End If
                                        If InStr(unformatted_matrix(i, 2), "PHASE|") > 0 And InStr(unformatted_matrix(i, 2), "END") > 0 Then
                                           'Resize Array
                                                sizeFlag1 = sizeFlag1 + 1
                                                ReDim Preserve phase2_mtx(3, sizeFlag1)
                                            'Set Unit
                                                lArray = Split(piTag(j), unitParseChar)
                                                phase1_mtx(0, sizeFlag) = lArray(unitPosition)
                                            'Batch ID
                                                lArray = Split(unformatted_matrix(i, 2), "|")
                                                phase2_mtx(1, sizeFlag1) = lArray(0)
                                            'PHASE
                                                lArray = Split(unformatted_matrix(i, 2), "|")
                                                lArray = Split(lArray(1), "\")
                                                lArray = Split(lArray(3), ":")
                                                phase2_mtx(2, sizeFlag1) = lArray(0)
                                            'Time
                                                phase2_mtx(3, sizeFlag1) = Format(unformatted_matrix(i, 1), "mm/dd/yyyy hh:mm:ss")
                                        End If
                                    End If
                                Next i
                            Else
                            End If
                    Next j
                    
                     'Combine Matrix
                        ReDim my_matrix(20, UBound(phase1_mtx, 2))
                        For i = 0 To UBound(phase1_mtx, 2)
                            For j = 0 To UBound(phase2_mtx, 2)
                                If phase1_mtx(0, i) = phase2_mtx(0, j) And phase1_mtx(1, i) = phase2_mtx(1, j) And phase1_mtx(2, i) = phase2_mtx(2, j) And phase1_mtx(3, i) < phase2_mtx(3, j) Then
                                    my_matrix(0, i) = phase1_mtx(0, i)
                                    my_matrix(1, i) = phase1_mtx(1, i)
                                    my_matrix(2, i) = phase1_mtx(2, i)
                                    my_matrix(3, i) = phase1_mtx(3, i)
                                    my_matrix(4, i) = phase2_mtx(3, j)
                                    Exit For
                                End If
                            Next j
                        Next i
                        
                        
                Else
                    unformatted_matrix = Application.Run("PICompDat", piTag, startTime, endTime, 9, "", "inside")
                        'Move to my_matrix
                            On Error Resume Next
                            If unformatted_matrix(1) <> "No more values:" Then
                            On Error Resume Next
                                For i = 1 To UBound(unformatted_matrix, 1)
                                    If unformatted_matrix(i, 1) <> " " And unformatted_matrix(1) <> "No more values:" Then
                                        If InStr(unformatted_matrix(i, 2), "PHASE|") > 0 And InStr(unformatted_matrix(i, 2), "START") > 0 Then
                                           'Resize Array
                                                sizeFlag = sizeFlag + 1
                                                ReDim Preserve phase1_mtx(3, sizeFlag)
                                            'Set Unit
                                                lArray = Split(piTag, unitParseChar)
                                                phase1_mtx(0, sizeFlag) = lArray(unitPosition)
                                            'Batch ID
                                                lArray = Split(unformatted_matrix(i, 2), "|")
                                                phase1_mtx(1, sizeFlag) = lArray(0)
                                            'PHASE
                                                lArray = Split(unformatted_matrix(i, 2), "|")
                                                lArray = Split(lArray(1), "\")
                                                lArray = Split(lArray(3), ":")
                                                phase1_mtx(2, sizeFlag) = lArray(0)
                                            'Time
                                                phase1_mtx(3, sizeFlag) = Format(unformatted_matrix(i, 1), "mm/dd/yyyy hh:mm:ss")
                                        End If
                                        If InStr(unformatted_matrix(i, 2), "PHASE|") > 0 And InStr(unformatted_matrix(i, 2), "END") > 0 Then
                                           'Resize Array
                                                sizeFlag1 = sizeFlag1 + 1
                                                ReDim Preserve phase2_mtx(3, sizeFlag1)
                                            'Set Unit
                                                lArray = Split(piTag, unitParseChar)
                                                phase1_mtx(0, sizeFlag) = lArray(unitPosition)
                                            'Batch ID
                                                lArray = Split(unformatted_matrix(i, 2), "|")
                                                phase2_mtx(1, sizeFlag1) = lArray(0)
                                            'PHASE
                                                lArray = Split(unformatted_matrix(i, 2), "|")
                                                lArray = Split(lArray(1), "\")
                                                lArray = Split(lArray(3), ":")
                                                phase2_mtx(2, sizeFlag1) = lArray(0)
                                            'Time
                                                phase2_mtx(3, sizeFlag1) = Format(unformatted_matrix(i, 1), "mm/dd/yyyy hh:mm:ss")
                                        End If
                                    End If
                                Next i
                                
                                'Combine Matrix
                                    ReDim my_matrix(20, UBound(phase1_mtx, 2))
                                    For i = 0 To UBound(phase1_mtx, 2)
                                        For j = 0 To UBound(phase2_mtx, 2)
                                            If phase1_mtx(0, i) = phase2_mtx(0, j) And phase1_mtx(1, i) = phase2_mtx(1, j) And phase1_mtx(2, i) = phase2_mtx(2, j) And phase1_mtx(3, i) < phase2_mtx(3, j) Then
                                                my_matrix(0, i) = phase1_mtx(0, i)
                                                my_matrix(1, i) = phase1_mtx(1, i)
                                                my_matrix(2, i) = phase1_mtx(2, i)
                                                my_matrix(3, i) = phase1_mtx(3, i)
                                                my_matrix(4, i) = phase2_mtx(3, j)
                                                Exit For
                                            End If
                                        Next j
                                    Next i
                            Else
                            End If
                End If

        dataMatrix = my_matrix
    End Sub




Public Sub Get_PcsPhaseAtTime()
        'SET PAGE CHARACTERISTICS
            Application.Calculation = xlCalculationManual
            Application.ScreenUpdating = False
        'DECLARE VARIABLE
            Dim substrings() As String
            Dim unformatted_matrix, my_matrix, phase1_mtx, phase2_mtx, batch_mtx, task_mtx As Variant
            Dim i, j As Integer
            Dim LString As String
            Dim lArray() As String
            Dim procedural_modal() As String
            Dim task_tag As String
            Dim sizeFlag, sizeFlag1 As Integer
            Dim firstPass, toggleStart As Boolean
            startTime = CDate(CDbl(CDate(endTime)) - 0.5)
            endTime = CDate(endTime)
    
            'DEFINE VARIABLES

                unformatted_matrix = Application.Run("PICompDat", piTag, startTime, endTime, 9, "", "inside")
                'Move to my_matrix
                    On Error Resume Next
                    If unformatted_matrix(1) <> "No more values:" Then
                    On Error Resume Next
                            If unformatted_matrix(i, 1) <> " " And unformatted_matrix(1) <> "No more values:" Then
                                If (InStr(unformatted_matrix(i, 2), "PHASE|") > 0 Or InStr(unformatted_matrix(i, 2), "PHASESTATE|") > 0) And InStr(unformatted_matrix(i, 2), "START") > 0 Then

                                    'Batch ID
                                        lArray = Split(unformatted_matrix(UBound(unformatted_matrix, 1), 2), "|")
                                        timedBatchId = lArray(0)
                                    'PHASE
                                        lArray = Split(unformatted_matrix(UBound(unformatted_matrix, 1), 2), "|")
                                        lArray = Split(lArray(1), "\")
                                        lArray = Split(lArray(3), ":")
                                        timedPhaseId = lArray(0)
                                End If
                            End If
                    End If

    End Sub


Public Sub GetPhase(phaseState As String)
    'Data Matrix
        'From PI
        '01-JAN-21 12:00:00  myPhaseName|PHASE|START||
        '01-JAN-21 12:00:00  myPhaseName|PHASE|END||
    'SET PAGE CHARACTERISTICS
        Application.Calculation = xlCalculationManual
        Application.ScreenUpdating = False
    'DECLARE VARIABLE
        Dim substrings() As String
        Dim unformatted_matrix, my_matrix, phase1_mtx, phase2_mtx, batch_mtx, task_mtx As Variant
        Dim i, j As Integer
        Dim LString As String
        Dim lArray() As String
        Dim procedural_modal() As String
        Dim task_tag As String
        Dim sizeFlag, sizeFlag1 As Integer
        Dim firstPass, toggleStart As Boolean
        startTime = CDate(startTime)
        endTime = CDate(endTime)

        'DEFINE VARIABLES
            ReDim my_matrix(20, 0)
            ReDim phase1_mtx(3, 0)
            ReDim phase2_mtx(3, 0)
            sizeFlag = -1
            sizeFlag1 = -1
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
                                        If InStr(unformatted_matrix(i, 2), phaseState) > 0 And InStr(unformatted_matrix(i, 2), "START") > 0 Then
                                           'Resize Array
                                                sizeFlag = sizeFlag + 1
                                                ReDim Preserve phase1_mtx(3, sizeFlag)
                                            'Set Unit
                                                lArray = Split(piTag(j), unitParseChar)
                                                phase1_mtx(0, sizeFlag) = lArray(unitPosition)
                                            'PHASE
                                                lArray = Split(unformatted_matrix(i, 2), "|")
                                                phase1_mtx(1, sizeFlag) = lArray(0)
                                            'Time
                                                phase1_mtx(2, sizeFlag) = Format(unformatted_matrix(i, 1), "mm/dd/yyyy hh:mm:ss")
                                        End If
                                        If InStr(unformatted_matrix(i, 2), phaseState) > 0 And InStr(unformatted_matrix(i, 2), "END") > 0 Then
                                           'Resize Array
                                                sizeFlag1 = sizeFlag1 + 1
                                                ReDim Preserve phase2_mtx(3, sizeFlag1)
                                            'Set Unit
                                                lArray = Split(piTag(j), unitParseChar)
                                                phase2_mtx(0, sizeFlag1) = lArray(unitPosition)
                                            'PHASE
                                                lArray = Split(unformatted_matrix(i, 2), "|")
                                                phase2_mtx(1, sizeFlag1) = lArray(0)
                                            'Time
                                                phase2_mtx(2, sizeFlag1) = Format(unformatted_matrix(i, 1), "mm/dd/yyyy hh:mm:ss")
                                        End If
                                    End If
                                Next i
                            Else
                            End If
                    Next j
                    
                     'Combine Matrix
                        ReDim my_matrix(20, UBound(phase1_mtx, 2))
                        For i = 0 To UBound(phase1_mtx, 2)
                            For j = 0 To UBound(phase2_mtx, 2)
                                If phase1_mtx(0, i) = phase2_mtx(0, j) And phase1_mtx(1, i) = phase2_mtx(1, j) And phase1_mtx(2, i) < phase2_mtx(2, j) Then
                                    my_matrix(0, i) = phase1_mtx(0, i)
                                    my_matrix(1, i) = phase1_mtx(1, i)
                                    my_matrix(2, i) = phase1_mtx(2, i)
                                    my_matrix(3, i) = phase1_mtx(3, i)
                                    my_matrix(4, i) = phase2_mtx(3, j)
                                    Exit For
                                End If
                            Next j
                        Next i
                        
                        
                Else
                unformatted_matrix = Application.Run("PICompDat", piTag, startTime, endTime, 9, "", "inside")
                    'Move to my_matrix
                        On Error Resume Next
                        If unformatted_matrix(1) <> "No more values:" Then
                        On Error Resume Next
                            For i = 1 To UBound(unformatted_matrix, 1)
                                If unformatted_matrix(i, 1) <> " " And unformatted_matrix(1) <> "No more values:" Then
                                    If InStr(unformatted_matrix(i, 2), phaseState) > 0 And InStr(unformatted_matrix(i, 2), "START") > 0 Then
                                       'Resize Array
                                            sizeFlag = sizeFlag + 1
                                            ReDim Preserve phase1_mtx(3, sizeFlag)
                                        'Set Unit
                                            lArray = Split(piTag, unitParseChar)
                                            phase1_mtx(0, sizeFlag) = lArray(unitPosition)
                                        'PHASE
                                            lArray = Split(unformatted_matrix(i, 2), "|")
                                            phase1_mtx(1, sizeFlag) = lArray(0)
                                        'Time
                                            phase1_mtx(2, sizeFlag) = Format(unformatted_matrix(i, 1), "mm/dd/yyyy hh:mm:ss")
                                    End If
                                    If InStr(unformatted_matrix(i, 2), phaseState) > 0 And InStr(unformatted_matrix(i, 2), "END") > 0 Then
                                       'Resize Array
                                            sizeFlag1 = sizeFlag1 + 1
                                            ReDim Preserve phase2_mtx(3, sizeFlag1)
                                        'Set Unit
                                            lArray = Split(piTag, unitParseChar)
                                            phase2_mtx(0, sizeFlag1) = lArray(unitPosition)
                                        'PHASE
                                            lArray = Split(unformatted_matrix(i, 2), "|")
                                            phase2_mtx(1, sizeFlag1) = lArray(0)
                                        'Time
                                            phase2_mtx(2, sizeFlag1) = Format(unformatted_matrix(i, 1), "mm/dd/yyyy hh:mm:ss")
                                    End If
                                End If
                            Next i
                            
                            'Combine Matrix
                                ReDim my_matrix(20, UBound(phase1_mtx, 2))
                                For i = 0 To UBound(phase1_mtx, 2)
                                    For j = 0 To UBound(phase2_mtx, 2)
                                        If phase1_mtx(0, i) = phase2_mtx(0, j) And phase1_mtx(1, i) = phase2_mtx(1, j) And phase1_mtx(2, i) < phase2_mtx(2, j) Then
                                            my_matrix(0, i) = phase1_mtx(0, i)
                                            my_matrix(1, i) = phase1_mtx(1, i)
                                            my_matrix(2, i) = phase1_mtx(2, i)
                                            my_matrix(3, i) = phase2_mtx(2, j)
                                            Exit For
                                        End If
                                    Next j
                                Next i
                        Else
                        End If
            End If

    dataMatrix = my_matrix
End Sub


