Attribute VB_Name = "GAMod"
Option Base 1 ' Start Array from 1
Option Explicit


Public Type Individual
Genome() As Integer 'Genome which holds information
Fitness As Double 'Fitness of Individual
MadeBy As String ' Made By
End Type

Public Type Population
NumOfIndivid As Integer 'Number of individuals
Individuals() As Individual 'Individuals
Parents() As Individual 'Parents
MaxFitness As Double 'Maximum fitness
NotifyWhenFitExceeds As Double 'Notify the user if fitness exceeds this value
GenomeLen As Integer 'Length of Genome
DiedIndivid() As Integer 'Individuals who have died coz they had low fitness value
NoOfDied As Integer 'Present number of died individuals
ProbMut As Double 'Mutation probability per cent
ProbCross As Double 'CrossOver probability per cent
StopEvolution As Boolean ' Check if we have to stop evolution
Generation As Integer ' Generation number
FitLim As Long 'Fitness limit below which all will be killed
BestSoFar As Individual ' Best Individual so far
WorstSoFar As Individual ' Worst Individual so far
End Type

Global PopuMain As Population ' Global population variable

Sub BuildPopu(Popu As Integer, LenghtOfGenome As Integer, MaxFit As Double, NotifyExceed As Double, Mut As Double, Cross As Double)
Dim i, j As Integer
PopuMain.Generation = 1 ' 1st genration
PopuMain.NumOfIndivid = Popu ' Population
PopuMain.GenomeLen = LenghtOfGenome 'Length of Genome
PopuMain.MaxFitness = MaxFit ' Maximum fitness
PopuMain.NotifyWhenFitExceeds = NotifyExceed
PopuMain.ProbMut = Mut
PopuMain.ProbCross = Cross
PopuMain.StopEvolution = False
ReDim PopuMain.Individuals(PopuMain.NumOfIndivid) As Individual
ReDim PopuMain.DiedIndivid(PopuMain.NumOfIndivid) As Integer
ReDim PopuMain.BestSoFar.Genome(PopuMain.GenomeLen)
ReDim PopuMain.WorstSoFar.Genome(PopuMain.GenomeLen)
For i = 1 To PopuMain.NumOfIndivid ' To generate random genome for every individual
    DoEvents
    ReDim PopuMain.Individuals(i).Genome(PopuMain.GenomeLen)
    PopuMain.Individuals(i).MadeBy = "Initial Population"
    For j = 1 To PopuMain.GenomeLen 'To generate random num between 0 and 9
        DoEvents
        PopuMain.Individuals(i).Genome(j) = RndNum(0, 9)
    Next j
Next i
End Sub

Public Function RndNum(Min As Long, Max As Long) As Double
'generates a random integer between the supplied values of Min and Max
    Randomize
    RndNum = Int((Max - Min + 1) * Rnd + Min)
End Function

Sub Evolve()
Do While Not PopuMain.StopEvolution = True
DoEvents
Dim i As Integer
For i = 1 To PopuMain.NumOfIndivid
    DoEvents
    FitnessTest PopuMain.Individuals(i) ' Assign a fitness value to each individual
    If PopuMain.Individuals(i).Fitness >= PopuMain.NotifyWhenFitExceeds Then 'Check if this individual is fit enough to provide a solution
        SolutionFound PopuMain.Individuals(i) ' Yes, solution is found
        If PopuMain.StopEvolution = True Then Exit Sub
    End If
    Debug.Print PopuMain.Individuals(i).Fitness
Next i
KillAllWorst
If PopuMain.NoOfDied <= Int((PopuMain.NumOfIndivid / 100) * 7) Then
    Mutate (True) ' Mutate all the 50% population
    KillAllWorst
ElseIf PopuMain.NoOfDied = PopuMain.NumOfIndivid Then ' All die
    Mutate True
    KillAllWorst
End If
'ElseIf RndNum(1, 100) < PopuMain.ProbCross Then
    KillAllWorst ' Kill all bad ones
    SelectParents 'Select parents
    CrossOver 'Cross
    If RndNum(1, 100) < PopuMain.ProbMut Then
        Mutate (False) ' Mutate
    End If
'Else
    'KillAllWorst
    'SelectParents
    'Reproduction
    'CrossOver
    'Mutate (True)
'End If
PopuMain.NoOfDied = 0 ' Reset all died individ
PopuMain.Generation = PopuMain.Generation + 1
Loop
End Sub


Sub KillAllWorst() ' Kill all the worst individuals
' The fitness limit is calculated as under
' FitLim = Max Fitness which can be acquired by individual - (MaxFitness of Individual - MinFitness Of individual)
PopuMain.NoOfDied = 0
Dim MaxFitnessVal As Double 'Max fitness in current population
Dim MinFitnessVal As Double 'Min fitness in current population
Dim SumFitness As Double ' Total sum of fitness
Dim i As Integer
Dim BestGenome, WorstGenome As String
Dim WorstIndex, BestIndex As Long
If PopuMain.Generation = 1 Then
MinFitnessVal = 0 'A small number so that maximum fitness can be recognized
MinFitnessVal = 9999 'A large num so that lowest fitness can be recognized
Else
MinFitnessVal = PopuMain.WorstSoFar.Fitness
MaxFitnessVal = PopuMain.BestSoFar.Fitness
End If
'Debug.Print "Difference:" & vbCrLf
For i = 1 To PopuMain.NumOfIndivid ' Calculte maximum and minimum fitness
    DoEvents
    If PopuMain.Individuals(i).Fitness > MaxFitnessVal Then
        MaxFitnessVal = PopuMain.Individuals(i).Fitness
        BestIndex = i
    End If
    If PopuMain.Individuals(i).Fitness < MinFitnessVal Then
        MinFitnessVal = PopuMain.Individuals(i).Fitness
        WorstIndex = i
    End If
    SumFitness = SumFitness + PopuMain.Individuals(i).Fitness
Next i
If WorstIndex <> Empty Then
For i = 1 To PopuMain.GenomeLen
    PopuMain.WorstSoFar.Genome(i) = PopuMain.Individuals(WorstIndex).Genome(i)
Next i
PopuMain.WorstSoFar.MadeBy = PopuMain.Individuals(WorstIndex).MadeBy
PopuMain.WorstSoFar.Fitness = PopuMain.Individuals(WorstIndex).Fitness
End If

For i = 1 To PopuMain.GenomeLen
    WorstGenome = WorstGenome & PopuMain.WorstSoFar.Genome(i)
Next i


If BestIndex <> Empty Then
PopuMain.BestSoFar.Fitness = PopuMain.Individuals(BestIndex).Fitness
PopuMain.BestSoFar.MadeBy = PopuMain.Individuals(BestIndex).MadeBy
For i = 1 To PopuMain.GenomeLen
    PopuMain.BestSoFar.Genome(i) = PopuMain.Individuals(BestIndex).Genome(i)
Next i
End If
For i = 1 To PopuMain.GenomeLen
    BestGenome = BestGenome & PopuMain.BestSoFar.Genome(i)
Next i

Call frmGA.PrntTextBest(CStr(PopuMain.BestSoFar.Fitness), CStr(BestGenome), PopuMain.BestSoFar.MadeBy)
Call frmGA.PrntTextWorst(CStr(PopuMain.WorstSoFar.Fitness), CStr(WorstGenome), PopuMain.WorstSoFar.MadeBy)
PopuMain.FitLim = SumFitness / PopuMain.NumOfIndivid
For i = 1 To PopuMain.NumOfIndivid ' Calculte maximum and minimum fitness
    DoEvents
    If PopuMain.Individuals(i).Fitness < PopuMain.FitLim Then ' Check if it is below limit
        PopuMain.NoOfDied = PopuMain.NoOfDied + 1 'Increase no of deaths
        PopuMain.DiedIndivid(PopuMain.NoOfDied) = i 'Individual is died
    End If
Next i
End Sub

Sub Reproduction()
Dim i, j As Integer
'Reproduce the best individuals According to thier fitness
For i = 1 To Int(PopuMain.NoOfDied / 2) 'Reproduce 50% only
    For j = 1 To PopuMain.GenomeLen
        PopuMain.Individuals(PopuMain.DiedIndivid(i)).Genome(j) = PopuMain.Parents(i).Genome(j) ' Copy the existing parent
    Next j
Next i
PopuMain.NoOfDied = PopuMain.NoOfDied - Int(PopuMain.NoOfDied / 2)
End Sub


Sub SelectParents() ' Select parents to produce childern who will replace those who have died
' The more fit the individual is the more is the chances that he will become a parent
Dim i, j, OffProduce, ParIndex, Index, Percent, Parent1, Parent2 As Integer
Dim MaxFit, MaxFitLim As Long
Percent = Int((PopuMain.NoOfDied / 100) * 50) ' 60 percent chance of selecting random parent
ReDim PopuMain.Parents(PopuMain.NoOfDied) As Individual
Index = 0
For i = 1 To Percent
    DoEvents
    Index = Index + 1
    Parent1 = RndNum(1, CLng(PopuMain.NumOfIndivid)) ' Select a random parent
    'Do
    'Parent2 = RndNum(1, CLng(PopuMain.NumOfIndivid)) ' Select a random parent
    'Loop While PopuMain.Individuals(Parent1).Fitness = PopuMain.Individuals(Parent2).Fitness
    'If PopuMain.Individuals(Parent2).Fitness < PopuMain.Individuals(Parent1).Fitness Then
    PopuMain.Parents(i) = PopuMain.Individuals(Parent1)
    'Else
    'PopuMain.Parents(i) = PopuMain.Individuals(Parent2)
    'End If
Next i
MaxFitLim = PopuMain.MaxFitness 'Maximum Fitness Level
MaxFit = 0 'Low number to catch highest fitness
i = Percent
Do ' Another 50 percent
MaxFit = 0
    For j = 1 To PopuMain.NumOfIndivid
        If PopuMain.Individuals(j).Fitness > MaxFit Then
            If PopuMain.Individuals(j).Fitness < MaxFitLim Then
                MaxFit = PopuMain.Individuals(j).Fitness
                ParIndex = j
            End If
        End If
    Next j
    OffProduce = Int(MaxFit / PopuMain.FitLim) ' Offsprings
    If OffProduce > Int((PopuMain.NoOfDied - i)) Then  'More children then remaining loop
        OffProduce = (PopuMain.NoOfDied - i)
    End If
    If OffProduce = 0 Then
    OffProduce = 1
    End If
    MaxFitLim = MaxFit ' To select next best individual
    j = 0
    Do
    i = i + 1
    PopuMain.Parents(i) = PopuMain.Individuals(ParIndex)
    j = j + 1
    Loop Until j = OffProduce
Loop Until i = PopuMain.NoOfDied
End Sub

Sub CrossOver() 'Parents produce new children to replace who have died
Dim CrossPos, i, j As Integer ' Random CrossOver Position
For i = 1 To PopuMain.NoOfDied
    DoEvents
    If i Mod 2 = 1 And i <> PopuMain.NoOfDied Then ' Only odd coz we have to select 2 parents
        CrossPos = RndNum(1, PopuMain.GenomeLen - 1)
        For j = 1 To PopuMain.GenomeLen
            DoEvents
            If j <= CrossPos Then
                PopuMain.Individuals(PopuMain.DiedIndivid(i)).Genome(j) = PopuMain.Parents(i).Genome(j)
                PopuMain.Individuals(PopuMain.DiedIndivid(i + 1)).Genome(j) = PopuMain.Parents(i + 1).Genome(j)
            Else
                PopuMain.Individuals(PopuMain.DiedIndivid(i)).Genome(j) = PopuMain.Parents(i + 1).Genome(j)
                PopuMain.Individuals(PopuMain.DiedIndivid(i + 1)).Genome(j) = PopuMain.Parents(i).Genome(j)
            End If
        Next j
    PopuMain.Individuals(PopuMain.DiedIndivid(i)).MadeBy = "CrossOver"
    PopuMain.Individuals(PopuMain.DiedIndivid(i + 1)).MadeBy = "CrossOver"
    End If
Next i
End Sub

Sub Mutate(MutWorst As Boolean) ' Mutation which will occur in all worst or one good
Dim IndividToMutate, i As Integer ' Individual to mutate
Dim RandMutPos, RandMutPos1 As Integer 'Random postion of mutation
Dim RandNum, RandNum1 As Integer 'Rand number
If MutWorst = True Then
For i = 1 To Int(PopuMain.NumOfIndivid / 3) ' Mutate 50%
    DoEvents
    RandMutPos = RndNum(1, CLng(PopuMain.GenomeLen))
    RandNum = RndNum(0, 9)
    Do
    RandMutPos1 = RndNum(1, CLng(PopuMain.GenomeLen))
    RandNum1 = RndNum(0, 9)
    Loop While (RandMutPos = RandMutPos1)
    PopuMain.Individuals(i).Genome(RandMutPos) = RandNum
    PopuMain.Individuals(i).Genome(RandMutPos1) = RandNum1
    PopuMain.Individuals(i).MadeBy = "Mutation of All Individuals"
    Next i
Else
    RandMutPos = RndNum(1, CLng(PopuMain.GenomeLen))
    RandNum = RndNum(0, 9)
    IndividToMutate = RndNum(1, CLng(PopuMain.NumOfIndivid))
    PopuMain.Individuals(IndividToMutate).Genome(RandMutPos) = RandNum
    PopuMain.Individuals(IndividToMutate).MadeBy = "Mutation of Single Individual"

End If
End Sub
