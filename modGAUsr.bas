Attribute VB_Name = "UsrMod"
Global Allow As Long 'Fitness which is allowed to be displayed
Global NumToFind As Integer ' The number which we have to find
Global Y_OFF As Integer
Global Y_OFF_E As Integer
Global X_OFF As Integer
Global Y_OFF_old As Integer
Global X_OFF_old As Integer
Public Sub FitnessTest(ByRef Individ As Individual) ' Do a fitness test
'It is upto you how you manipulate the data and assign the fitness value
'We have to find x, y and z such that x^2+y^2+z^2=90
Dim Valu As Integer

For i = 1 To PopuMain.GenomeLen
    Valu = Valu + Individ.Genome(i) ^ 2
    strr = strr & Individ.Genome(i)
Next i
If Valu > NumToFind Then
Individ.Fitness = (NumToFind / Valu) * PopuMain.MaxFitness
ElseIf Valu < NumToFind Then
Individ.Fitness = (Valu / NumToFind) * PopuMain.MaxFitness
Else
Individ.Fitness = PopuMain.MaxFitness
End If
Debug.Print strr
X_OFF = X_OFF + 1
Y_OFF = Y_OFF_E - Int(Individ.Fitness)
DrwLine
End Sub

Public Sub SolutionFound(Individ As Individual)
' This sub is called when fitness of an individual
' exceeds the value given by you as NotifyWhenFitExceeds
'Debug.Print "Solution Found : Acurracy =  " & Individ.Fitness & "% Generation = " & PopuMain.Generation & " Numbers  =  " & Individ.Genome(1) & "^2 + " & Individ.Genome(2) & "^2 + " & Individ.Genome(3) & "^2  = 100 "
If Individ.Fitness >= Allow Then
    MsgBox "Solution Found : Acurracy = " & Individ.Fitness & "% ; Made By : " & Individ.MadeBy & " ; Generation = " & PopuMain.Generation & " ; Numbers  =  " & Individ.Genome(1) & "^2 + " & Individ.Genome(2) & "^2 + " & Individ.Genome(3) & "^2 + " & Individ.Genome(4) & "^2 + " & Individ.Genome(5) & "^2  = " & NumToFind
    PopuMain.StopEvolution = True
End If
End Sub

Public Sub DrwLine()
frmGA.Line (X_OFF_old, Y_OFF_old)-(X_OFF, Y_OFF), vbBlack
X_OFF_old = X_OFF
Y_OFF_old = Y_OFF
If X_OFF > frmGA.ScaleWidth Then
X_OFF = 0
X_OFF_old = 0
frmGA.Cls
End If
End Sub
