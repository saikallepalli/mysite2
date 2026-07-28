import pandas as pd
df = pd.read_pickle('df.pkl')
col = df.columns[14]  # data check action points - alumni verification calls happen here
kws = ['shown as placed','shown working','confirmed working','not joined','doubtful','physical visit','confirmed he is working','confirmed on their joining','left job','working in','found working']
for idx,row in df.iterrows():
    val = row[col]
    if pd.isna(val): continue
    s = str(val)
    if any(k.lower() in s.lower() for k in ['verif','placed','doubtful','not joined','left job','confirmed']):
        print('---', row['Centre Name'], '|', row['Manager'])
        print(s[:600])
        print()
