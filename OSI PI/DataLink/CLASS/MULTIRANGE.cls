VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MULTIRANGE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Covington OSI PI Library
'Written 24AUG2020 by Nicholas Stom
'Class is used for a range of floating values

Option Explicit
' These are properties
    Public maxValue As Double
    Public minValue As Double
    Public timeAtMax As Date
    Public timeAtMin As Date
    Public startTime As Date
    Public endTime As Date
    Public piTag As Variant
    Public sampleTime As String
    Public dataMatrix As Variant

' These are methods
    Public Sub Get_MultiRange()
        'SET PAGE CHARACTERISTICS
            Application.Calculation = xlCalculationManual
            Application.ScreenUpdating = False
        'DECLARE VARIABLE
            Dim unformatted_matrix, my_matrix, batch_mtx, task_mtx As Variant
            Dim firstPass As Boolean
            Dim sizeFlag, i, j As Integer
            startTime = CDate(startTime)
            endTime = CDate(endTime)

            'DEFINE VARIABLES
                firstPass = False
                ReDim my_matrix(20, 0)
                sizeFlag = -1
                i = 1
                If IsArray(piTag) = True Then
                    For j = 0 To UBound(piTag)
                        unformatted_matrix = Application.Run("PISampDat", piTag(j), startTime, endTime, sampleTime, 1, "")
                        On Error Resume Next
                        If unformatted_matrix(1) <> "No more values:" Then
                        On Error Resume Next
                            If firstPass = False Then
                                For i = 1 To UBound(unformatted_matrix, 1)
                                    If unformatted_matrix(i, 1) <> " " And unformatted_matrix(1) <> "No more values:" Then

                                        sizeFlag = sizeFlag + 1
                                        ReDim Preserve my_matrix(20, sizeFlag)
                                        my_matrix(0, sizeFlag) = Format(unformatted_matrix(i, 1), "mm/dd/yyyy hh:mm:ss")
                                        my_matrix(j + 1, sizeFlag) = unformatted_matrix(i, 2)

                                    End If
                                Next i
                                
                            Else
                                sizeFlag = 0
                                For i = 1 To UBound(unformatted_matrix, 1)
                                    If unformatted_matrix(i, 1) <> " " And unformatted_matrix(1) <> "No more values:" Then
                                        my_matrix(j + 1, sizeFlag) = unformatted_matrix(i, 2)
                                        sizeFlag = sizeFlag + 1
                                    End If
                                Next i
                                firstPass = True
                            End If
                        Else
                        End If
                    firstPass = True
                    Next j

                Else
                End If
        dataMatrix = my_matrix
    End Sub
 
