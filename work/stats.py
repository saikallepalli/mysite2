import pandas as pd, re
df = pd.read_pickle('df.pkl')

sections = {
 'Data Check': list(range(6,14)),
 'Infra Check': list(range(15,23)),
 'Process Check': list(range(24,46)),
 'Aspirant Interaction': list(range(47,58)),
}

rows=[]
for sec, idxs in sections.items():
    for i in idxs:
        col = df.columns[i]
        label = re.sub(r'^\d?\.?\s*(Data Check|Infra Check|Process Check|Aspirant Interaction)?\s*\(approx\.[^)]*\)\s*\[?','',col).rstrip(']')
        vc = df[col].astype(str).value_counts()
        total = vc.sum()
        na = vc.get('Not Applicable',0)
        aw = vc.get('All is Well',0)
        ni = vc.get('Need Improvement',0)
        applicable = aw+ni
        pct_ni = round(100*ni/applicable,1) if applicable>0 else None
        rows.append({'Section':sec,'Item':label,'All is Well':aw,'Need Improvement':ni,'Not Applicable':na,'%NI (of applicable)':pct_ni,'Applicable N':applicable})

res = pd.DataFrame(rows)
pd.set_option('display.max_colwidth', 100)
pd.set_option('display.width', 220)
res_sorted = res.sort_values('%NI (of applicable)', ascending=False)
print(res_sorted.to_string(index=False))
res.to_csv('item_stats.csv', index=False)
