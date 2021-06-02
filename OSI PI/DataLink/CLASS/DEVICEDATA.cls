VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DEVICEDATA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Covington OSI PI Library
'Written 24AUG2020 by Nicholas Stom
'Class is used for a range of floating values

Option Explicit
' These are properties
    Public startTime As Date
    Public endTime As Date
    Public piTag As String
    Public cycledOn As Integer
    Public cycledOff As Integer
    Public actuations As Integer
    Public timeOn As Date
    Public timeOff As Date
    Public dataMatrix As Variant

' These are methods
    Public Sub Get_DeviceData()
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
            startTime = CDate(startTime)
            endTime = CDate(endTime)
            cycledOn = 0
            cycledOff = 0
            actuations = 0
            timeOn = 0
            timeOff = 0
    
    
            'DEFINE VARIABLES
                firstPass = False
                ReDim my_matrix(1, 0)
                sizeFlag = -1
                i = 1

                unformatted_matrix = Application.Run("PICompDat", piTag, startTime, endTime, 9, "", "inside")
            'Move to my_matrix
                On Error Resume Next
                If unformatted_matrix(1) <> "No more values:" Then
                On Error Resume Next
                    For i = 1 To UBound(unformatted_matrix, 1)
                        If unformatted_matrix(i, 1) <> " " And unformatted_matrix(1) <> "No more values:" Then
                            
                            If firstPass = True Then
                                If unformatted_matrix(i, 2) = "OPEN" And unformatted_matrix(i - 1, 2) = "CLOSED" Then
                                    onTimeStart = unformatted_matrix(i, 1)
                                    offTimeEnd = unformatted_matrix(i, 1)
                                    cycledOn = cycledOn + 1
                                    actuations = actuations + 1
                                    generalOffTime = generalOffTime + (offTimeEnd - offTimeStart)
                                ElseIf unformatted_matrix(i, 2) = "CLOSED" And unformatted_matrix(i - 1, 2) = "OPEN" Then
                                    offTimeStart = unformatted_matrix(i, 1)
                                    onTimeEnd = unformatted_matrix(i, 1)
                                    cycledOff = cycledOff + 1
                                    actuations = actuations + 1
                                    generalOnTime = generalOnTime + (onTimeEnd - onTimeStart)
                                End If
                            End If
                            

                            
                            If firstPass = False Then
                                If unformatted_matrix(i, 2) = "OPEN" Then
                                    onTimeStart = unformatted_matrix(i, 1)
                                ElseIf unformatted_matrix(i, 2) = "CLOSED" Then
                                    offTimeStart = unformatted_matrix(i, 1)
                                End If
                                firstPass = True
                            End If
                            
                            
                            
                        End If
                    Next i
                Else
                End If
        
         timeOn = Format(CDate(generalOnTime), "h:m:s")
         timeOff = Format(((endTime - startTime) - generalOnTime), "hhhhh:mm:sss")
    End Sub
    
    Public Sub Get_DeviceDataAtTime()
        'SET PAGE CHARACTERISTICS
            Application.Calculation = xlCalculationManual
            Application.ScreenUpdating = False
        'DECLARE VARIABLE
            Dim unformatted_matrix, my_matrix, batch_mtx, task_mtx As Variant

            'DEFINE VARIABLES


            unformatted_matrix = Application.Run("PIArcVal", piTag, startTime, 0, "", "auto")
            dataMatrix = unformatted_matrix

    End Sub
 
