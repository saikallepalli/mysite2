import pandas as pd
df = pd.read_pickle('df.pkl')
action_cols = [df.columns[14], df.columns[46]]
count=0
for idx,row in df.iterrows():
    for c in action_cols:
        val = row[c]
        if pd.isna(val): continue
        s = str(val)
        if 'dob' in s.lower() or 'under age' in s.lower() or 'underage' in s.lower() or '18 year' in s.lower() or 'below age' in s.lower():
            count+=1
            print('---', row['Centre Name'],'|',row['Manager'])
            print(s[:250])
            print()
print('TOTAL', count)
