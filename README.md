# Stock Analysis using VBA

# Overview of the Project
The purpose of this project is to automate the calculations, using VBA Macros, of the Daily Volume and Yearly Return of stocks for the companies selected for this dataset. For educational puroses, the worksheet *All Stocks Analysis Pivot Table* provides the same calculations done with the macros, but performed with a Pivot Table and the *VLOOKUP* function.

# Results

## Stock Results for 2017
The results obtained for 2017 indicate that it was a good year for the selected companies. *TERP* was the only company with a negative yearly return. *DQ* and *ENPH*, FSLR* and *SEDG* surpassed the 100% return.
2018 Stocks Results
:-------------------:
![](Resources/Stocks_2017.png)
## Performance for 2017
The nested loops macro took about 0.85 seconds to complete while the refactored version took about 0.15 seconds, that is a **567%** increase in performance.
Nested Loop Macro                | Refactored Macro
:-------------------------------:|:-------------------------------------:
![](Resources/Original_2017.png) | ![](Resources/VBA_Challenge_2017.png) 

## Stock Results for 2018
The results obtained for 2017 indicate that it was a bad year for the selected companies with *JKS* and *DQ* having a return of under **-60%**. *ENPH* and *RUN* were the only company with a positive yearly return. Interestingly, most companies had a greater volume in 2018 than in 2017, which indicates that even though there were more transactions, this companies lost credibility.
2018 Stocks Results
:-------------------:
![](Resources/Stocks_2018.png)
## Performance for 2018
The nested loops macro took about 0.8 seconds to complete while the refactored version took about 0.16 seconds, that is a 500% increase in performance.
Nested Loop Macro                | Refactored Macro
:-------------------------------:|:-------------------------------------:
![](Resources/Original_2018.png) | ![](Resources/VBA_Challenge_2018.png) 


# Summary

## Advantages and Disadvantages of Refactoring
Refactoring has many advantages and disadvantages. Here are some:

### Advantages
- Improve code readability.
- Simplification of logics.
- Reduced complexity reduction.
- Better modeling of classes and functions.
- Helps discover hidden or dormant bugs.
- Improves mantainability.

### Disadvantages
- Time is spent away from development.
- Management usually doesn't care about code mantainability.
- May introduce new bugs.

## Advantages and Disadvantages applied to the Refactoring of the Original VBA Macro

### Advantages
- The avoidance of nested loops makes it easier to follow the logic of the program.
- Using the fact that the dataset is ordered by *TICK* and *DATE*, the refactored code uses a simple loop instead of a nested loop which reduces the complexity from O(n*m) to O(n) with n being the number of rows and m the nmber of ticks.
- With a simpler logic, the code is much easier to debug.
- The refactored code was over 5 times faster than the original one.

### Disadvantages
- Time Consuming.
- Uses a lot of assumptions about the dataset structure.
