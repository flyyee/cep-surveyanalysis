# cep-surveyanalysis
__Outcome:__
_Analyses survey results and presents analysis, in the form of console text and visual graphs._


__Task:__
Analysing CS Feedback Survey

Problem Statement
From student feedback on their Year 1 CS experience, your Computing teacher wants to find out answers to a few questions:

Predicting the demand of Y2 CEP: How many of them are interested to choose CEP in Year 2?
Consolidated picture of how they perceive the 9 weeks CS course.

These are some possible ways to present the findings:

Rudimentary:
For Numeric responses: 
Calculate average score for the questions with numeric values
For Categorical responses: 
Calculate percentage of students who chose a particular response
Intermediate:
Calculates other measures of central tendency [refer to http://www.abs.gov.au/websitedbs/a3121120.nsf/home/statistical+language+-+measures+of+central+tendency] 

E.g. Median, Mode
Advanced: 
The survey may incorporate lesser / more questions over different years. Your program allows the ability to process lesser / more questions. Perhaps allowing a config file to specify which column contains what kind of data, e.g. general surveyee info, numeric, categorical data
Presenting your analysis using visuals like charts / graphs. (needs external modules like matplotlib https://realpython.com/python-matplotlib-guide/ ) 
Summarizing free-responses text (needs external modules like sumy https://github.com/miso-belica/sumy) 

Input file
responses.xlsx

Output
Report of findings. 
Could be console output, csv, Excel, Word or  saved Matplotlib image file. Though console output is least preferred. 

Files you need:
Link: https://tinyurl.com/y2csanalysis 
responses.xlsx
