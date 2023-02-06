This repo was initially created to help automate a tedious task for a friend interning at a finance firm. The task they needed to complete was to find a finance article for each individual trade that was executed during the past 9 months. Then the articles' URLs needed to be added to a [trading log](https://github.com/agarcia50S/SideProjects/blob/master/redacted_log.xlsx). Each article needed to satisty 2 criteria: it mentions the traded stock at least once and it's publication date is before the trade's date.

I was given the [trading log which was an Excel Workbook that contained 12 spreadsheets](https://github.com/agarcia50S/SideProjects/blob/master/redacted_log.xlsx), each containing the trades made during a given month. The data in each table, however, was not properly organized, making any attempt to access the data awkward and convoluted.

Therefore, I set out to complete 2 main tasks in this project:

1) Clean and reformat the data into the standard table structure.
2) Build a bot to to retrieve the url of a finance article from the internet based on a given ticker name and a time frame.

[redacted_log.xlsx](https://github.com/agarcia50S/SideProjects/blob/master/redacted_log.xlsx) is the redacted version of the trading log and [file_redaction.py](https://github.com/agarcia50S/SideProjects/blob/master/file_redaction.py) is the code used to redact it.
