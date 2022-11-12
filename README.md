WORK IN PROGRESS: refactor code into class objects and to better handle errors in raw data

This repo was initially created to help friend interning at a company automate a task. The task was to automate the process of searching for a finance article about 
a given stock within a given time range. I was given a ghastly structured spreadsheet that was being used as a trade log. After explaining why
the current structuring of the data was flawed I was asked to also clean/reformat the data. 

Therefore, I set out to complete 2 main tasks in this project:

1) Clean and reformat the data into a basic flat file structure (i.e. 2-D structure composed of columns/rows)
2) Build a bot to retrieve to retrieve the url to a relavent finance article from the web based on a given ticker name and date.


[redacted_log.xlsx](https://github.com/agarcia50S/SideProjects/blob/master/redacted_log.xlsx) is a redacted version of the trade log and [file_redaction.py](https://github.com/agarcia50S/SideProjects/blob/master/file_redaction.py) is the code used to redact it.
