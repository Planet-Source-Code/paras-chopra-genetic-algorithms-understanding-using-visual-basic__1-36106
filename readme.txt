A new Genetic Algorithm (with example) - Works brilliantly - Update

Here comes the Genetic Algorithm library which is little different from other algorithms. First, it uses real values instead of binary ones. Second, there is very little problem specific code needs to be typed. Third,unlike other algorithms either new offspring will replace the bad ones or mutation occurs in all the bad ones.
 
You only have to do 4 things
1. Build population using buildpop sub.
2. Call Evolve method.
3. Build a sub which will evaluate fitness.
4. Build a sub which is called when solution is found.
 
The example consists of - library(GAMod),  Module(UsrMod) and Form(frmGA)
The purpose of this example is to take a number and resolves it into form like v^2+w^2+x^2+y^2+z^2, where v, w, x, y and z are between 0 and 9.
 
Changes:
1. A little more nice GUI
2. A progress graph
3. Information about best and worst chromosomes so far
4. Information about the process from which a specified chromosome is made
5. And most important: change in algorithm
 
Note: The module is customised a litlle bit. The non customised module is in directory "GA_NonCustomised", which you can use in your applications. It may be having bugs.
 
Solve complex problems with this algorithm. And have fun. Please VOTE! and COMMENTS are appreciated

Made By Paras Chopra
CEO, NaramCheez
http://naramcheez.netfirms.com
paraschopra@lycos.com

P.S.:If you want a tutorial on GA please tell me. I can write one.