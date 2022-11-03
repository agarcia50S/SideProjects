#%%
import pandas as pd
import xlrd

def get_sheet_names_xls(filepath):
    xls = xlrd.open_workbook(filepath, on_demand=True)
    return xls.sheet_names()

def clean_df(df, cols, parse_on, how_dropna='all'):
    df = df.loc[:, :parse_on].drop(columns=cols).dropna(how=how_dropna)
    return df.dropna(axis='columns', how=how_dropna).reset_index(drop=True)

def row_as_tuple(df, cols, c):
    return (df[col][c] for col in cols)

def month_col(df, col):
    temp = []
    date = ''
    for i in df[col]:
        if type(i) is pd.Timestamp:
            date = i.strftime('%B %#d, %Y')
        elif type(i) is str:
            date = i
        temp.append(date)
    return temp

if __name__ == '__main__':
    path = 'C:/Users/agarc/AbbyCode/data/'
    all_dfs = {}

    for df in get_sheet_names_xls(path+'trade_confirm_log.xls')[:2]:
        pre_clean = pd.read_excel(path+'trade_confirm_log.xls', sheet_name=df)
        post_clean = pre_clean.drop(columns=['Signed by DG']).dropna(how='all').reset_index(drop=True)
        post_clean['Trade Date'] = month_col(post_clean, 'Trade Date')
    
        all_dfs[df] = post_clean
    #%%
    for df in get_sheet_names_xls(path+'trade_confirm_log.xls')[2:]:
        og = pd.read_excel(path+'trade_confirm_log.xls', sheet_name=df)

        # saving col to re-add to df later
        check = [i for i in og['Signed by DG'] if type(i) is str][0]
        # prev_len = len(og)
        cleaned = clean_df(og, ['Signed by DG'], 'REASON for TRADE')
        # making col same len as cleaned
        # checks = checks[:len(checks) - (prev_len - len(og))]
        
        cleaned['Trade Date'] = month_col(cleaned, 'Trade Date')

        c = 0
        temp_dict = {i:[] for i in cleaned.columns.values}

        tickers = cleaned.pop('COMPANY (TICKER)')

        try:
            while c < len(cleaned):
                v0, v1, v2, v3, = row_as_tuple(cleaned, cleaned.columns.values, c)
                cur_tick = tickers[c]
                if type(cur_tick) is str:
                    for tick in cur_tick.split(','):
                        temp_dict['Trade Date'].append(v0)
                        temp_dict['Entity/ Individual'].append(v1)
                        temp_dict['Action'].append(v2)
                        temp_dict['REASON for TRADE'].append(v3)
                        temp_dict['COMPANY (TICKER)'].append(tick)
                else:
                    temp_dict['Trade Date'].append(v0)
                    temp_dict['Entity/ Individual'].append(v1)
                    temp_dict['Action'].append(v2)
                    temp_dict['REASON for TRADE'].append(v3)
                    temp_dict['COMPANY (TICKER)'].append(cur_tick)
                    
                c += 1
        except Exception as error:
            print(error)
            print(tickers[c])
            break
                        
        final = pd.DataFrame(temp_dict)
        all_dfs[df] = final
    #%%
    print(all_dfs['January'])

    for v in all_dfs.values():
        print(type(v))
    # %%
    with pd.ExcelWriter("data/result.xlsx") as writer:
        for k, v in all_dfs.items():
                # use to_excel function and specify the sheet_name and index
                # to store the dataframe in specified sheet
                v.insert(loc=2, 
                        column='Signed by DG', 
                        value=check
                )
                v.to_excel(writer, sheet_name=k, index=False)
# %%
