<div align="center">

## Genetic Algorithms : Understanding using Visual Basic\.


</div>

### Description

This tutorial is written with the a strong aim. It's main purpose is to make you understand Genetic Algorithms(GA). This tutorial is written by using examples from the 'Genetic Algorithm using Real Numbers' written by me. It is written in complementary to my code. This tutorial was written because in my sight there was not a single tutorial with both Genetic Algorithms and Visual Basic in it. I have written this tutorial to help a current VB programmer understand and implement GA with ease.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2002-06-18 09:34:54
**By**             |[Paras Chopra](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/paras-chopra.md)
**Level**          |Intermediate
**User Rating**    |4.9 (93 globes from 19 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Genetic\_Al971886212002\.zip](https://github.com/Planet-Source-Code/paras-chopra-genetic-algorithms-understanding-using-visual-basic__1-36106/archive/master.zip)





### Source Code

```
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>Genetic Algorithms : A brief tutorial.</title>
</head>
<body>
<center><h1>Genetic Algorithms : Understanding using Visual Basic.</h1></center>
<p><big><big>1. About this tutorial</big></big>
<p><blockquote>This tutorial is written with the a strong aim. It's main purpose
is to make you understand Genetic Algorithms(GA). This tutorial is written by using
examples from the 'Genetic Algorithm using Real Numbers' written by me. It is written in
complementary to my code.</blockquote>
<p><big><big>2. Why this tutorial?</big></big>
<p><blockquote>This tutorial was written because in my sight there was not a single
tutorial with both Genetic Algorithms and Visual Basic in it. I have written this tutorial
to help a current VB programmer understand and implement GA with ease.</blockquote>
<p><big><big>3. Introduction</big></big>
<p><blockquote>Genetic Algorithms, What are they? Well ,there is not a single strict definition of
GA. Every author gives his own definition. So, I am not going to increase the cluttering but rather
give a definition by another author.
<p><center><i> Genetic Algorithms are programs that simulate the logic of Darwinian selection, if you understand how populations accumulate differences over time due to the environmental conditions acting as a selective breeding mechanism then you understand GAs. Put another way, understanding a GA means understanding the simple, iterative processes that underpin evolutionary change.
</i></center>
<p>A little bit confusing, isn't it. So, should I expand it to a simple definition. Yes, I should.
GA is a algorithm which makes it easy to search a large <i>search space</i>. For example, if we have a number
and we want to find its largest divisor. The problem is too easy if the number is small. But the
complexity increases as the number increases. Here GAs are used. GAs used Darwinian selection. In
<i>'The Origin of Species'</i>, Darwin stated that from a group of individuals the best will survive.
By implementing this Darwinian selection to the problem only the best solutions will remain narrowing the
search space.</blockquote>
<p><big><big>4. Where GAs can be used?</big></big>
<p><blockquote>GAs can be used where optimization is needed. I mean that where there are large
solutions to the problem but we have to find the best one. Like we can use GAs in finding best moves in chess,
mathematical problems, financial problems and in many more areas.</blockquote>
<p><big><big>5. Are there any disadvantages of GAs?</big></big>
<p><blockquote>Yes, there are few disadvantages. But, remember there are more advantages than
disadvantages. Disadvantages<ul><li>GAs are very slow.</li><li>
They cannot always find the exact solution but they always find best solution.</li></ul></blockquote>
<p><big><big>6. Explanation of terms</big></big>
<p><blockquote><i><big>Chromosome: </big></i> A set of genes. Chromosome contains the solution in form of genes.
<br><i><big>Gene: </big></i> A part of chromosome. A gene contains a part of solution. It determines the solution. E.g. 16743 is a chromosome and 1, 6, 7, 4 and 3 are its genes.
<br><i><big>Individual: </big></i> Same as chromosome.
<br><i><big>Population: </big></i> No of individuals present with same length of chromosome.
<br><i><big>Fitness: </big></i> Fitness is the value assigned to an individual. It is based on how far
or close a individual is from the solution. Greater the fitness value better the solution it contains.
<br><i><big>Fitness function: </big></i>Fitness function is a function which assigns fitness value to the individual.
It is problem specific.
<br><i><big>Breeding: </big></i>Taking two fit individuals and intermingling there chromosome to create
new two individuals.
<br><i><big>Mutation: </big></i>Changing a random gene in an individual.
<br><i><big>Selection: </big></i>Selecting individuals for creating the next generation.
</blockquote>
<p><big><big>7. General Algorithm of GA.</big></big>
<p><blockquote>The algorithm is almost same in most of the applications only fitness functions
are different to different problems. The general algorithm is as follows :
<p><b>Start<br>
<blockquote>Generate initial population.
<br>Assign fitness function to all individuals.
<p>DO UNTIL best solution is found
<blockquote>Select individuals from current generation
<br>Create new offsprings with mutation and/or breeding
<br>Compute new fitness for all individuals
<br>Kill all the unfit individuals to give space to new offsprings
<br>Check if best solution is found</blockquote>
END</blockquote>END</b></blockquote></blockquote>
<p><big><big>8. Explanation of Genetic Algorithm this using Visual Basic</big></big>
<p><blockquote>Here is the explanation of GA coded by me in Visual Basic. My algorithm has some differences
from the general algorithm.</blockquote>
<p><blockquote><b><big>Define an individual</big></b>
<blockquote><pre>Public Type Individual
Genome() As Integer 'Genome which holds information
Fitness As Double 'Fitness of Individual
MadeBy As String ' Made By</blockquote>
End Type</pre>Here Genome is the array of integers. My algorithm uses integers instead of
binary numbers. Integers are easy to handle and probably more efficient in GAs. It's fitness contains
the fitness value. And MadeBy holds the information about the process from which it is made.
<p><b><big>Define population</big></b>
<blockquote><pre>Public Type Population
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
End Type</blockquote>
</pre>The code above is self explanatory.
<p><b><big>Build the population</big></b>
<blockquote>
Sub BuildPopu(Popu As Integer, LenghtOfGenome As Integer, MaxFit As Double, NotifyExceed As Double, Mut As Double, Cross As Double)
</blockquote>'Popu' is the number of individuals. 'LenghtOfGenome' is the length of chromosome
'MaxFit' is the maximum fitness which can be acquired by an individual. It is generally
100. 'NotifyExceed' is the range which tells the algorithm to notify the user when fitness of any individual
goes beyond this level. It is generally 99. 'Mut' and 'Cross' are probabilities.
<p><b><big>Evolve the population</big></b>
<p>Evolution of the population is defined by following algorithm<blockquote><b>
DO UNTIL StopEvolution = True
<blockquote>Assign Fitness to each individual
<br>Notify the user if a solution is found
<br>Kill all the worst individuals
<p>IF less then 30% of population is dead Or All the population is dead then
<blockquote>Mutate all 33% of the population
<br>Kill all the worst individuals
</blockquote>END
<p>Kill all the worst individuals
<br>Select the parents
<br>Start breeding
<br>Mutate a random individual if probability allows.
</blockquote>LOOP</b></blockquote>
In the above algorithm the algorithm mutates 33% of population if less or all individuals are dying
because in both the situations the necessary evolution does not take place. And crossing is done every
time because without crossing new generation cannot be made.
<p><b><big>Selection of parents</big></b>
<p>33% of the parents are selected on the basis of their fitness i.e. fitter the parent more the children he
will have and another 33% of parents will be selected randomly.
<p><b><big>CrossOver</big></b>
<p>Take two individuals from the parents list. And then take a random crossover point.
Interchange the genes to produce two new individuals. For example let the two parents
be 1234 and 5678 and the random crossover point be 2 then two new individuals will be
1278 and 5634.
<p><b><big>Mutation</big></b>
<p>Take any random individual and take a random point. Change the gene
on that point with another random value.</blockquote>
<p><big><big>9. How come my algorithm is little different from others?</big></big>
<p>There are several differences. Some of them are below and you will discover
other differences while you study the code.<ul><li>It uses real numbers instead of binary numbers.
<li>CrossOver is done in every generation
<li>33% of individuals are mutated if death rate falls down below 30%.
<li>Written in Visual Basic
<li>It can be applied to variety of problems very easily
</ul>
<p><big><big>9. About the author</big></big>
<p><blockquote>My name is Paras Chopra. I am mad about AI. This tutorial is written in
complementary to my code. I am interested in Neural Networks, NeuroGenetics, Genetic Programming
and other fields of AI. I am also interested in starting a site
dedicated to works of AI written in Visual Basic. If I get positive feedback about the site.
I will start it. Please give comments and If possible please VOTE!
</body>
</html>
```

