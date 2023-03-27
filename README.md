# VBA-Challenge-Submission
The code analyses data on a workbook both the 'sample' and the 'actual'. On sample as mentioned in the assignment my code runs in less than 2 minutes. I have created a button on cutom ribbon located on top left hand corner of the spreadsheet.

I wrote a pseudo code as mentioned by James and have then referred to various websites and resources to learn how to perform a specific action that I want to perform. Learnt a lot of new commands and different ways of doing things. Some of then resources that I have reffered to are but not limited to YouTube, www.automateexcel.com, excelchamps.com, www.educba.com, www.wallstreetmojo.com.

My code is divided in following parts:

1. Declaring Variables
2. 3x For Loops to collect data, process, analyse and then print as per the requirement. Funtionality of each loop is as below:
    Loop#1 - Loops through different Sheets of a Workbook. 
             Following actions are also performed in Loop#1 once we get the output from Loop#2
             a. Analyses Percent Change and Total Stock Volume
             b. Prints values: Greatest Percentage Increase, Greatest Percentage Decrease and Greatest Total Volume
    Loop#2 - Loops through different ticker symbols and assigns a specific ticker to process to Loop#3. 
             Once we get the output from Loop#3, following actions are performed:
             a. Prints values: Ticker Symbol, Yearly Change, Percent Change, Total Stock Volume
             b. Conditional formatting applied to Yearly Change and Percent Change
    Loop#3 - Collects Open Price, Close Price, Total Volume for a specific ticker and passes on to Loop#2
3. Formatting - I have added additional functionality to update the font type, size, alignment and width for the range (A:Q)

I have uploaded the sample spreadsheet with the results as the actual data file was too large but the code works exactly the same on the actual data file.

Please let me know if you have any questions.
