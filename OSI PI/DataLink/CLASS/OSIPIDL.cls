VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OSIPIDL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'OSI PI DataLink Library
'Written by Nicholas Stom

Option Explicit
' These are properties
    Public startTime As Date
    Public endTime As Date
    Public sampleTime As Variant
    Public piTag As Variant
    Public dataMatrix As Variant



' These are methods
    'Returns Device Data in Matrix
    Public Sub GetDeviceData()
        'SET PAGE CHARACTERISTICS
            Application.Calculation = xlCalculationManual
            Application.ScreenUpdating = False
        'DECLARE VARIABLE
            Dim unformatted_matrix, my_matrix, partialMtx As Variant
            Dim currentValue, currentPrevValue, currentTime, currentPrevTime As Variant
            Dim firstPass As Boolean
            Dim sizeFlag, I, j As Integer
            Dim onTimeStart As Date
            Dim onTimeEnd As Date
            Dim offTimeStart As Date
            Dim offTimeEnd As Date
            Dim generalOnTime As Double
            Dim generalOffTime As Double
            startTime = CDate(startTime)
            endTime = CDate(endTime)
            ReDim my_matrix(80, 80)
    
            'DEFINE VARIABLES
                firstPass = True
                sizeFlag = -1
                I = 1
                currentPrevValue = "Blank"
                If IsNull(sampleTime) Then
                    unformatted_matrix = Application.Run("PICompDat", piTag, startTime, endTime, 9, "", "inside")
                Else
                   unformatted_matrix = Application.Run("PISampDat", piTag, startTime, endTime, sampleTime, 1, "")
                End If
                
            'Move to my_matrix
                On Error Resume Next
                If unformatted_matrix(1) <> "No more values:" Then
                On Error Resume Next
                    For I = 1 To UBound(unformatted_matrix, 1)
                        If unformatted_matrix(I, 1) <> "Resize to show all values" Then
                        If unformatted_matrix(I, 1) <> " " And unformatted_matrix(1) <> "No more values:" Then
                            If firstPass = True Then
                                currentValue = unformatted_matrix(I, 2)
                                currentTime = unformatted_matrix(I, 1)
                                currentPrevTime = unformatted_matrix(I, 1)
                                firstPass = False
                            End If
                            
                            currentValue = unformatted_matrix(I, 2)
                            currentTime = unformatted_matrix(I, 1)
                                
                            If currentPrevValue <> currentValue Then
                                For j = 0 To UBound(my_matrix, 2)
                                    If (my_matrix(0, j) = currentValue) Then
                                        my_matrix(1, j) = my_matrix(1, j) + 1
                                        Exit For
                                    ElseIf (my_matrix(0, j) = "") Then
                                        my_matrix(0, j) = currentValue
                                        my_matrix(1, j) = my_matrix(1, j) + 1
                                        Exit For
                                    End If
                                    
                                Next j
                                
                                For j = 0 To UBound(my_matrix, 2)
                                    If (my_matrix(0, j) = currentPrevValue) Then
                                        my_matrix(2, j) = my_matrix(2, j) + (currentTime - currentPrevTime)
                                        Exit For
                                    End If
                                    
                                Next j
                                currentPrevValue = currentValue
                                currentPrevTime = unformatted_matrix(I, 1)
                            End If
                            

                            
                            
                        End If
                        End If
                    Next I
                    For j = 0 To UBound(my_matrix, 2)
                        If (my_matrix(0, j) <> "") Then
                           my_matrix(2, j) = Round((my_matrix(2, j) * 24), 3) & " hours"
                        End If
                    Next j
                Else
                End If
                
        dataMatrix = my_matrix
    End Sub

    Public Sub GetDataAtTime()
        'SET PAGE CHARACTERISTICS
            Application.Calculation = xlCalculationManual
            Application.ScreenUpdating = False
        'DECLARE VARIABLE
            Dim unformatted_matrix As Variant

        'DEFINE VARIABLES
            unformatted_matrix = Application.Run("PIArcVal", piTag, startTime, 0, "", "auto")
            dataMatrix = unformatted_matrix
    End Sub
