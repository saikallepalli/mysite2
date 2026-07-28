import pandas as pd
df = pd.read_pickle('df.pkl')
action_cols = [df.columns[14], df.columns[23], df.columns[46], df.columns[58]]
col14 = df.columns[14]

print("=== VERACITY / PLACEMENT RED FLAGS ===")
for idx,row in df.iterrows():
    val = row[col14]
    if pd.isna(val): continue
    s = str(val)
    if any(k.lower() in s.lower() for k in ['shown as placed','shown working','confirmed working','not joined','doubtful','left job','confirmed he is working','confirmed on their joining','found working and']):
        print('---', row['Centre Name'], '|', row['Manager'])
        print(s[:400])
        print()

print("=== DOB / AGE FLAGS ===")
for idx,row in df.iterrows():
    for c in [df.columns[14], df.columns[46]]:
        val = row[c]
        if pd.isna(val): continue
        s = str(val)
        if 'dob' in s.lower() or 'under age' in s.lower() or 'underage' in s.lower() or 'below age' in s.lower():
            print('---', row['Centre Name'], '|', row['Manager'])
            print(s[:200])
            print()
