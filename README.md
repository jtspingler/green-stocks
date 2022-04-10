# green-stocks

Using VBA to analyze stocks over various time periods

## Overview of Project

In this project we practiced writing VBA in order to measure the returns of a handful of stocks over the years 2017 and 2018. In order to do to do this we used loops to go through our arrays that we created and then compile all returns for said stock. After compiling these, we formatted the outputs using some VBA functions that we went over in our modules last week. Once this was created, we were tasked with executing on some refactored code, which is a cleaned-up version of the first code we wrote in the stocks analysis. We measured the performance by using a timer to time how long the macros took to execute. 

### Purpose

The purpose of this project was to not only practice some of the code that we learned last week, but also to understand how valuable refactored code can be. I anticipate putting this lesson to practice as I continue to improve on my technical abilities in the coming weeks. Another unlock due to our learnings in this module is the automation of formatting. In some of my previous roles I have spent many long hours manually formatting worksheets. Automating this saves much time and creates some consistency as well.
 
## Analysis and Challenges

 The findings of my analysis can be found [here](https://github.com/jtspingler/green-stocks), as well as presented below:

We created buttons and assigned them to the macros we created which adds some nice functionality as well as aesthetics.
![This is an image](https://github.com/jtspingler/green-stocks/blob/9f60928e82792d3acb053853bc878d5beb8d4fbb/All%20Stocks%20blank.PNG) 

Here are the results of our macros for 2017 and 2018 before we refactored the code: 

![This is an image](https://github.com/jtspingler/green-stocks/blob/9f60928e82792d3acb053853bc878d5beb8d4fbb/All%20Stocks%202017%20timer.PNG)

![This is an image](https://github.com/jtspingler/green-stocks/blob/9f60928e82792d3acb053853bc878d5beb8d4fbb/All%20Stocks%202018%20timer.PNG)

### Outcomes with Refactored Code

![This is an image](https://github.com/jtspingler/green-stocks/blob/9f60928e82792d3acb053853bc878d5beb8d4fbb/Refactored%20All%20Stocks%202017%20timer%20formatted.PNG)

![This is an image](https://github.com/jtspingler/green-stocks/blob/9f60928e82792d3acb053853bc878d5beb8d4fbb/All%20Stocks%202018%20timer.PNG)

### Challenges and Difficulties Encountered

I definitely struggled creating the for loops through our multiple variables but the module helped me work through my issues. One particular area that I struggled through was changing the names of the sheets that we needed to activate in order to set our years to "yearvalue". To debug this I used the debugger in VBA and that highlighted the areas where I was making mistakes.

## Results

- What are two conclusions you can draw about the Outcomes based on the Refactored vs normal macro times?

2017 Normal: 55101.54 seconds
2018 Normal: 55185.86 seconds
2017 Refactored: .1210938 seconds
2018 Refactored: .1289063 seconds

It is clear that the refactored code is preference here, as it allows our code to run much more efficiently. This is especially important to keep in mind when working on larger projects

- What can you conclude about the Outcomes based on stock performances?

2017 was by far a much better year for investing, as 11/12 of the stocks we analyzed had positive returns. 2018 on the other hand was almost a complete reversal, with 10/12 stocks had negative returns.

- What are the advantages or disadvantages of refactoring code?

Clearly the advantages are that the code runs much more efficiently and is less cumbersome to read and write. Depending on how far along you are in a project, refactoring can be expensive and take up quite a bit of time. Further, I would be worried about introducing further bugs that may break something in our code.
