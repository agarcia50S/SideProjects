#%%
import pandas as pd
import random

def make_row_as_header(df, row):
    '''
    Takes DataFrame and replaces its current headers with
    the values of a given row
    RETURNS new DataFrame instance with new headers

    row -- index position of a given row
    '''
    output = df.rename(columns=df.iloc[row]).drop(df.index[row])
    return output

def redact_string(string, start=65, stop=90):
    '''
    Iterates over passed string, adding a randomly selected
    65-to-90 ASCII char to the list. List items are 
    concatenated into a single string.
    RETURNS string

    string -- str type object
    '''
    return ''.join([chr(random.randint(start, stop)) for _ in string])

def redact_ticker_col(col):
    '''
    Iterates over ticker col, spliting the multi-ticker strings into lists.
    Using a list comprehension, each ticker in a list is iteratively 
    passed into redact_string(). The resulting list is joined into a 
    string and appended into the output list.
    RETURNS list of redacted ticker values.

    col -- ticker column with uniquely formatted values
    '''
    redacted_col = []
    for val in col:
        # only redact strings
        if type(val) != str:
            redacted_col.append(val)
        else:
            # redacts 1 val in ticker col
            redacted_val = [redact_string(ticker) for ticker in val.split(',')]
            redacted_col.append(', '.join(redacted_val)) # append redacted val
    return redacted_col

def redact_names(col):
    '''
    via a dict comprehension, a list of the passed column's unique values 
    is iterated over, creating a map of names to redacted names. The map is
    applied to column using a .apply() and a lambda fnc.
    RETURNS a Series object

    col -- name column
    '''
    name_map = {name: redact_string(name) if type(name) == str else name for name in col.unique()}
    return col.apply(lambda x: name_map[x])

def redact_links(col):
    '''
    Iterates over passed column via list comprehension, adding redacted
    link if value is string type.

    col -- column with links as string type values 
    '''
    return ['https://'+redact_string(link, start=97, stop=122) if type(link) == str else link for link in col]

#%%
path = 'C:/Users/agarc/AbbyCode/data/'
og = pd.read_excel(path+'unedited_trade_log.xls', sheet_name='April')
og.head()

# %%
# storing row 0 for later use
row_0 = pd.DataFrame(og.iloc[[0]])

# remove row 0 for easier wrangling
work_df = og.drop(index=0)
work_df.head()

# %%

# replace confidential data with redacted data
work_df['Unnamed: 4'] = redact_ticker_col(work_df['Unnamed: 4'])
work_df['Unnamed: 1'] = redact_names(work_df['Unnamed: 1'])
work_df['Unnamed: 7'] = redact_links(work_df['Unnamed: 7'])
work_df.head()

# %%
# add row 0 back to main df
concated = pd.concat([row_0, work_df])

# name col headers 1-6 "nan" to match original file's table structure
new_headers = [None if 'Unnamed' in header else header for header in concated.columns]
concated.columns = new_headers

concated.head()

# %%
concated.to_excel(path + 'test_alpha.xlsx', index=False)
# %%
