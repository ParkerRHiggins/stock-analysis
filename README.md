# stock-analysis

## Overview of Project
### Background
Steve has asked me to help him analyze a green stock that his parents are interested in investing in.  The data that we’re analyzing is the annual return percentage over two years for 12 different green stocks including the green stock that Steve’s parents are particularly interested in to see which green stock is best investing in.  Steve has some experience working with Excel manually but utilizing Visual Basic for Applications, or VBA, we can write code that will automate our analysis for us.  
### Purpose
The purpose of this project was to analyze multiple green stocks efficiently utilizing VBA.  Once a working code was created that successfully analyzed multiple green stock return information for 2017 and 2018, we wanted to see if we could then modify the code to make it run more efficiently and ultimately decrease the codes run time.  We achieved this by refactoring the original code which resulted in a much quicker run time.

## Results
While refactoring the code it was important to decide which parts of the original code we can still use and copy over and areas that we can improve the code.  I was able improve the code by creating three new output arrays (tickerVolume, tickerStartingPrices and tickerEndingPrices) and a new tickIndex variable that we used to access the correct index across the arrays before looping through the data set.  This strategy allowed our code to run quicker and more efficiently.  Below are the original code and run times for yearValue verses and new refractured code and run times. 
### Original Code


### Refractured Code


## Summary
### Advantages and Disadvantages of Refactoring Code in General
Refactoring code has both advantages and disadvantages.  An advantage to refactoring code is that the code runs more efficiently and results in a quicker run time for the code.  Also, the code becomes more organized and easier to read which is important for debugging and when other users are trying to analyze your code.
A disadvantage to refracturing code is that your manipulated code that already runs correctly and risking making changes that could result in breaking the code.  This is a calculated risk when deciding to refractor code and why it’s always important to save new versions of the code with notes regularly utilized GitHub.
### Advantages and Disadvantages of the Original and Refactored VBA Script
When refactoring the stock analysis code for Steve, the biggest advantage was that the run time for running the macro decreased from about 0.5 seconds to about 0.1 seconds.  That’s a major increase in efficiency in the code and that difference can really add up over time if the code is run often.  
A disadvantage when refracturing the stock analysis code is that it took some time to figure out why I was receiving error messages in the code which required me to spend hours trying to figure out how to resolve the errors.  Nevertheless, it was worth refactoring the stock analysis code for Steve and a great learning experience.
