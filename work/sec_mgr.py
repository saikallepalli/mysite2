import pandas as pd
df = pd.read_pickle('df.pkl')
sections = {
 'Data Check': list(range(6,14)),
 'Infra Check': list(range(15,23)),
 'Process Check': list(range(24,46)),
 'Aspirant Interaction': list(range(47,58)),
}
# section aggregate
for sec, idxs in sections.items():
    cols = [df.columns[i] for i in idxs]
    sub = df[cols].astype(str)
    aw = (sub=='All is Well').sum().sum()
    ni = (sub=='Need Improvement').sum().sum()
    na = (sub=='Not Applicable').sum().sum()
    app = aw+ni
    print(sec, 'AllWell',aw,'NeedImp',ni,'NA',na, '%NI', round(100*ni/app,1))

print()
# manager wise: overall NI rate across all checklist cols, and NA usage
all_idx = sections['Data Check']+sections['Infra Check']+sections['Process Check']+sections['Aspirant Interaction']
cols = [df.columns[i] for i in all_idx]
mgr_rows=[]
for mgr, g in df.groupby('Manager'):
    sub = g[cols].astype(str)
    aw = (sub=='All is Well').sum().sum()
    ni = (sub=='Need Improvement').sum().sum()
    na = (sub=='Not Applicable').sum().sum()
    app = aw+ni
    mgr_rows.append({'Manager':mgr,'Visits':len(g),'AllWell':aw,'NeedImprovement':ni,'NA':na,
                      '%NI':round(100*ni/app,1) if app else None,
                      '%NA':round(100*na/(aw+ni+na),1)})
mgr_df = pd.DataFrame(mgr_rows).sort_values('%NI', ascending=False)
print(mgr_df.to_string(index=False))
mgr_df.to_csv('manager_stats.csv', index=False)
