import pandas as pd
import openpyxl


if __name__ == '__main__':
    wb = openpyxl.open('output.xlsx')
    sheet = wb.get_sheet_by_name('Sheet1')
    df = pd.DataFrame(sheet.values)
    result = pd.DataFrame(columns=["category",
                                   "man",
                                   "name",
                                   "mod",
                                   "color",
                                   "color_ru"])
    for i, row in df.iterrows():
        dict_to_df = {"category": row[1],
                      "man": row[2],
                      "name": row[3],
                      "mod": row[4],
                      "color": row[5],
                      "color_ru": row[6]
                      }
        if (row[3] in result['name'].unique()) and (row[4] in result['mod'].unique()) and (row[5] in result['color'].unique()):
            pass
            print(i)
        else:
            result = result.append(dict_to_df, ignore_index=True)
            print(dict_to_df)
writer = pd.ExcelWriter('output_2.xlsx')
result.to_excel(writer)
writer.save()