import pandas as pd
df = pd.read_pickle('df.pkl')
sections = {
 'Data Check': list(range(6,14)),
 'Infra Check': list(range(15,23)),
 'Process Check': list(range(24,46)),
 'Aspirant Interaction': list(range(47,58)),
}
all_idx = sections['Data Check']+sections['Infra Check']+sections['Process Check']+sections['Aspirant Interaction']
cols = [df.columns[i] for i in all_idx]
action_cols = [df.columns[14], df.columns[23], df.columns[46], df.columns[58]]

def wc(x):
    if pd.isna(x): return 0
    return len(str(x).split())
df['action_words'] = df[action_cols].apply(lambda r: sum(wc(v) for v in r), axis=1)

rows=[]
for mgr, g in df.groupby('Manager'):
    sub = g[cols].astype(str)
    aw=(sub=='All is Well').sum().sum(); ni=(sub=='Need Improvement').sum().sum(); na=(sub=='Not Applicable').sum().sum()
    app=aw+ni
    rows.append({'Manager':mgr,'Visits':len(g),'%NI':round(100*ni/app,1),'%NA':round(100*na/(aw+ni+na),1),'AvgWords':round(g['action_words'].mean())})
res = pd.DataFrame(rows).sort_values('Visits', ascending=False)
print(res.to_string(index=False))
