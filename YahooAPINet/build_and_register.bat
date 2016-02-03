@echo Buiding C# API
csc /target:library FinanceCS.cs

@echo Buiding VB.Net API
vbc /target:library FinanceVBNet.vb

@echo Buiding C++/CLI API
cl /clr:safe /LD FinanceCPP.cpp

@echo Registering C# API
regasm /codebase /tlb FinanceCS.dll

@echo Registering VB.Net API
regasm /codebase /tlb FinanceVBNet.dll

@echo Registering C++/CLI API
regasm /codebase /tlb FinanceCPP.dll