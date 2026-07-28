import pandas as pd
df = pd.read_pickle('df.pkl')
action_cols = [df.columns[14], df.columns[23], df.columns[46], df.columns[58]]
keywords = ['not joined','not working','doubtful','disciplinary','shown as placed','ghost','fraud','left job','not join','edited document','typo','under age','underage','below age','wrong DOB','DOB']
hits = []
for idx, row in df.iterrows():
    for c in action_cols:
        val = row[c]
        if pd.isna(val): continue
        val_s = str(val)
        for kw in keywords:
            if kw.lower() in val_s.lower():
                hits.append({'Centre': row['Centre Name'], 'Manager': row['Manager'], 'Keyword': kw, 'Snippet': val_s[:300]})
                break
        else:
            continue
        break
h = pd.DataFrame(hits)
print(len(h))
print(h[['Centre','Manager','Keyword']].to_string())
