VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DATARANGE"
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
    Public starttime As Date
    Public endtime As Date
    Public piTag As String
    Public dataMatrix As Variant

' These are methods
    Public Sub Get_DataRange()
        'SET PAGE CHARACTERISTICS
            Application.Calculation = xlCalculationManual
            Application.ScreenUpdating = False
        'DECLARE VARIABLE
            Dim unformatted_matrix, my_matrix, batch_mtx, task_mtx As Variant
            Dim firstPass As Boolean
            Dim sizeFlag, i As Integer
            starttime = CDate(starttime)
            endtime = CDate(endtime)
    
            'DEFINE VARIABLES
                firstPass = False
                ReDim my_matrix(1, 0)
                sizeFlag = -1
                i = 1
                unformatted_matrix = Application.Run("PISampDat", piTag, starttime, endtime, "00:00:05", 1, "")
            'Move to my_matrix
                On Error Resume Next
                If unformatted_matrix(1) <> "No more values:" Then
                On Error Resume Next
                    For i = 1 To UBound(unformatted_matrix, 1)
                        If unformatted_matrix(i, 1) <> " " And unformatted_matrix(1) <> "No more values:" Then
                            If firstPass = False Then
                                maxValue = unformatted_matrix(i, 2)
                                timeAtMax = Format(unformatted_matrix(i, 1), "mm/dd/yyyy hh:mm:ss")
                                minValue = unformatted_matrix(i, 2)
                                timeAtMin = Format(unformatted_matrix(i, 1), "mm/dd/yyyy hh:mm:ss")
                                firstPass = True
                            End If

                            sizeFlag = sizeFlag + 1
                            ReDim Preserve my_matrix(1, sizeFlag)
                            my_matrix(0, sizeFlag) = Format(unformatted_matrix(i, 1), "mm/dd/yyyy hh:mm:ss")
                            my_matrix(1, sizeFlag) = unformatted_matrix(i, 2)
                    
                            'Find Min And Max
                                If my_matrix(1, sizeFlag) > maxValue Then
                                    maxValue = my_matrix(1, sizeFlag)
                                    timeAtMax = Format(my_matrix(0, sizeFlag), "mm/dd/yyyy hh:mm:ss")
                                End If
                                If my_matrix(1, sizeFlag) < minValue Then
                                    minValue = my_matrix(1, sizeFlag)
                                    timeAtMin = Format(my_matrix(0, sizeFlag), "mm/dd/yyyy hh:mm:ss")
                                End If
                        End If
                    Next i
                Else
                End If
        dataMatrix = my_matrix
    End Sub
    
    Public Sub Get_DeviceDataCompressed()
        'SET PAGE CHARACTERISTICS
            Application.Calculation = xlCalculationManual
            Application.ScreenUpdating = False
        'DECLARE VARIABLE
            Dim unformatted_matrix, my_matrix, batch_mtx, task_mtx As Variant
            Dim firstPass As Boolean
            Dim sizeFlag, i As Integer
            Dim onTimeStart As Date
            Dim onTimeEnd As Date
            Dim offTimeStart As Date
            Dim offTimeEnd As Date
            Dim generalOnTime As Double
            Dim generalOffTime As Double
            starttime = CDate(starttime)
            endtime = CDate(endtime)
            cycledOn = 0
            cycledOff = 0
            actuations = 0
            timeOn = 0
            timeOff = 0
    
    
            'DEFINE VARIABLES
                firstPass = False
                ReDim my_matrix(20, 0)
                sizeFlag = -1
                i = 1

                unformatted_matrix = Application.Run("PICompDat", piTag, starttime, endtime, 9, "", "inside")
            'Move to my_matrix
                On Error Resume Next
                If unformatted_matrix(1) <> "No more values:" Then
                On Error Resume Next
                    For i = 1 To UBound(unformatted_matrix, 1)
                        If unformatted_matrix(i, 1) <> " " And unformatted_matrix(1) <> "No more values:" Then
                            'Resize Array
                                sizeFlag = sizeFlag + 1
                                ReDim Preserve my_matrix(20, sizeFlag)
                                my_matrix(0, i) = Format(unformatted_matrix(i, 1), "mm/dd/yyyy hh:mm:ss")
                                my_matrix(1, i) = unformatted_matrix(i, 2)
                        End If
                    Next i
                Else
                End If
        dataMatrix = my_matrix

    End Sub
 
