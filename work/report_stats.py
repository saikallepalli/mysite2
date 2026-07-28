import pandas as pd, re
df = pd.read_pickle('df.pkl')

sections = {
 'Data Check': list(range(6,14)),
 'Infra Check': list(range(15,23)),
 'Process Check': list(range(24,46)),
 'Aspirant Interaction': list(range(47,58)),
}

def clean_label(col):
    m = re.search(r'\[(.*)\]', col)
    return m.group(1).strip() if m else col.strip()

rows=[]
for sec, idxs in sections.items():
    for i in idxs:
        col = df.columns[i]
        label = clean_label(col)
        vc = df[col].astype(str).value_counts()
        aw = vc.get('All is Well',0); ni = vc.get('Need Improvement',0); na = vc.get('Not Applicable',0)
        app = aw+ni
        rows.append({'Section':sec,'Item':label,'AW':aw,'NI':ni,'NA':na,'App':app,'Pct': round(100*ni/app,1) if app else None})
res = pd.DataFrame(rows)
res_sorted = res.sort_values(['NI','Pct'], ascending=False)
pd.set_option('display.max_colwidth', 100); pd.set_option('display.width', 200)
print("TOP 12 WEAKEST ITEMS:")
print(res_sorted.head(12).to_string(index=False))
print()

# Section aggregate
print("SECTION AGGREGATE:")
for sec, idxs in sections.items():
    cols=[df.columns[i] for i in idxs]
    sub = df[cols].astype(str)
    aw=(sub=='All is Well').sum().sum(); ni=(sub=='Need Improvement').sum().sum(); na=(sub=='Not Applicable').sum().sum()
    app=aw+ni
    print(sec, 'AW',aw,'NI',ni,'NA',na,'%NI',round(100*ni/app,1))
