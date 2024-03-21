import pandas as pd
import numpy as np 
import json
df = pd.read_excel(r"D:\python\Credit_Rating_Details.xlsx")
df
sf = pd.DataFrame(df["Rating Details"])
sf
def extract_keys_values(row):
    try:
        json_data = json.loads(row)
        row_dict = {}
        for entry in json_data:
            for key, value in entry.items():
                if key in row_dict:
                    row_dict[key].append(value)
                else:
                    row_dict[key] = [value]
        return pd.Series(row_dict)
    except Exception as e:
        print(f"Error: {e}")
        return pd.Series({})


sf = sf['Rating Details'].apply(extract_keys_values)
sf = sf.applymap(lambda x: x[0] if isinstance(x, list) else x).fillna('')

print(sf)
pd.DataFrame(sf)
sf.head(10)
sf.isnull().sum()
sf.info()
row_with_date = sf[sf['date'] == "2024-07-03"]

pd.DataFrame(row_with_date)
sf.rename(columns={'orgNumber': 'CIN'}, inplace=True)

sf = pd.DataFrame(sf)
sf.head()
columns = sf.columns.tolist()
columns.remove('CIN')
columns.insert(0, 'CIN')
sf = sf[columns]
print(sf)
sf.head
sf = pd.DataFrame(sf)
sf
data = pd.read_excel(r"D:\python\Screener_CIN_Mapping.xlsx")
data
data_2 = pd.merge(sf, data, on="CIN" )
data_2 = data_2.drop(["companyName"], axis=1)
data_2.head(1)
data_2.rename(columns={'date': 'Date of Rating',
                        'agency': "Rating Agency", 
                        'amount':"Amount (Mn)",
                        'rating':"Rating", 
                        "instrument":"Instrument",
                        "ratingGrade":"Rating Grade",
                        "ratingStatus":"Rating Status",
                        'Screener Link':"Link"}, inplace=True)

data_2.head()
data_2.to_excel('output1.xlsx', index=False)
