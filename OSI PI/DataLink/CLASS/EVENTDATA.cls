VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EVENTDATA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Covington OSI PI Library
'Written 24AUG2020 by Nicholas Stom
'Class is used to Get PCS Events Run on Unit
'   The PiTag is a public variant which means it can be an array of multiple units

Option Explicit
' These are properties
    Public startTime As Date
    Public endTime As Date
    Public piTag As Variant
    Public dataMatrix As Variant

' These are methods
    Public Sub Get_Event()
        'SET PAGE CHARACTERISTICS
            Application.Calculation = xlCalculationManual
            Application.ScreenUpdating = False
        'DECLARE VARIABLE
            Dim substrings() As String
            Dim unformatted_matrix, my_matrix, phase1_mtx, phase2_mtx, batch_mtx, task_mtx As Variant
            Dim I, j As Integer
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
                                For I = 1 To UBound(unformatted_matrix, 1)
                                   If unformatted_matrix(I, 1) <> " " And unformatted_matrix(1) <> "No more values:" Then
                                        
                                           'Resize Array
                                                sizeFlag = sizeFlag + 1
                                                ReDim Preserve my_matrix(20, sizeFlag)
                                            'Set Unit
                                                lArray = Split(piTag(j), "_")
                                                If InStr(lArray(2), "CL") > 0 Then
                                                    my_matrix(0, sizeFlag) = lArray(2)
                                                ElseIf InStr(lArray(3), "CL") > 0 Then
                                                    my_matrix(0, sizeFlag) = lArray(3)
                                                ElseIf InStr(lArray(4), "CL") > 0 Then
                                                    lArray = Split(lArray(4), ".")
                                                    my_matrix(0, sizeFlag) = lArray(0)
                                                End If
                                            
                                            'PHASE
                                                my_matrix(1, sizeFlag) = "No Phase Identified"
                                                On Error Resume Next
                                                lArray = Split(unformatted_matrix(I, 2), "|")
                                                On Error Resume Next
                                                lArray = Split(lArray(1), ":")
                                                If InStr(lArray(0), "PH_") > 0 Then
                                                    my_matrix(1, sizeFlag) = lArray(0)
                                                End If
                                                On Error Resume Next
                                            'Event
                                                my_matrix(2, sizeFlag) = unformatted_matrix(I, 2)
                                            'Time
                                                my_matrix(3, sizeFlag) = Format(unformatted_matrix(I, 1), "mm/dd/yyyy hh:mm:ss")
                                    
                                    End If
                                Next I
                            Else
                            End If
                    Next j
                    
                        
                        
                Else
                    unformatted_matrix = Application.Run("PICompDat", piTag, startTime, endTime, 9, "", "inside")
                        'Move to my_matrix
                            On Error Resume Next
                            If unformatted_matrix(1) <> "No more values:" Then
                            On Error Resume Next
                                On Error GoTo ErrorHandler
                                If UBound(unformatted_matrix, 2) > 1 Then
                                    On Error GoTo ErrorHandler
                                    For I = 1 To UBound(unformatted_matrix, 1)
                                        If unformatted_matrix(I, 1) <> " " And unformatted_matrix(1) <> "No more values:" Then
                                            If unformatted_matrix(I, 1) <> "Resize to show all values" Then
                                                'Resize Array
                                                    sizeFlag = sizeFlag + 1
                                                    ReDim Preserve my_matrix(20, sizeFlag)
                                                'Set Unit
                                                    lArray = Split(piTag, "_")
                                                     If InStr(lArray(2), "CL") > 0 Then
                                                        my_matrix(0, sizeFlag) = lArray(2)
                                                    ElseIf InStr(lArray(3), "CL") > 0 Then
                                                        my_matrix(0, sizeFlag) = lArray(3)
                                                    ElseIf InStr(lArray(4), "CL") > 0 Then
                                                        lArray = Split(lArray(4), ".")
                                                        my_matrix(0, sizeFlag) = lArray(0)
                                                    End If
                                            
                                                'PHASE
                                                    my_matrix(1, sizeFlag) = "NA"
                                                    On Error Resume Next
                                                    lArray = Split(unformatted_matrix(I, 2), "|")
                                                    On Error Resume Next
                                                    lArray = Split(lArray(1), ":")
                                                    If InStr(lArray(0), "PH_") > 0 Then
                                                        my_matrix(1, sizeFlag) = lArray(0)
                                                    End If
                                                    On Error Resume Next
                                                'Event
                                                    my_matrix(2, sizeFlag) = unformatted_matrix(I, 2)
                                                'Time
                                                    my_matrix(3, sizeFlag) = Format(unformatted_matrix(I, 1), "mm/dd/yyyy hh:mm:ss")
                                                End If
                                            
                                    
                                        End If
                                    Next I
                                Else
                                    sizeFlag = sizeFlag + 1
                                    ReDim Preserve my_matrix(20, sizeFlag)
                                    'Event
                                        my_matrix(2, sizeFlag) = unformatted_matrix(1)
                                    'Time
                                        my_matrix(3, sizeFlag) = Format(unformatted_matrix(2), "mm/dd/yyyy hh:mm:ss")
                                End If
                                
                                
                            Else
                            End If
                End If

        dataMatrix = my_matrix
    Exit Sub
ErrorHandler:
        sizeFlag = sizeFlag + 1
        ReDim Preserve my_matrix(20, sizeFlag)
        'Event
            my_matrix(2, sizeFlag) = unformatted_matrix(2)
        'Time
            my_matrix(3, sizeFlag) = Format(unformatted_matrix(1), "mm/dd/yyyy hh:mm:ss")
        dataMatrix = my_matrix
                                
    End Sub







