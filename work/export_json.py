import pandas as pd, re, json
df = pd.read_pickle('df.pkl')

sections = {
 'Data Check': list(range(6,14)),
 'Infra Check': list(range(15,23)),
 'Process Check': list(range(24,46)),
 'Aspirant Interaction': list(range(47,58)),
}

def clean_label(col, sec_prefix_num):
    # strip leading numbering/section name and trailing bracket
    m = re.search(r'\[(.*)\]', col)
    if m:
        return m.group(1).strip()
    return col.strip()

items_meta = []  # list of {section, label, col_index}
for sec, idxs in sections.items():
    for i in idxs:
        col = df.columns[i]
        label = clean_label(col, sec)
        items_meta.append({'section': sec, 'label': label, 'idx': i})

visits = []
for _, row in df.iterrows():
    rec = {
        'manager': row['Manager'],
        'centre': row['Centre Name'],
        'program': row['Program Name'],
        'date': str(row['Visit Date'])[:10],
        'items': [],
        'q_bottleneck': (row[df.columns[59]] if pd.notna(row[df.columns[59]]) else ''),
        'q_placement': (row[df.columns[60]] if pd.notna(row[df.columns[60]]) else ''),
        'q_batch_ontime': (row[df.columns[61]] if pd.notna(row[df.columns[61]]) else ''),
    }
    for im in items_meta:
        val = row[df.columns[im['idx']]]
        val = str(val) if pd.notna(val) else 'Not Applicable'
        rec['items'].append({'section': im['section'], 'label': im['label'], 'status': val})
    visits.append(rec)

out = {
    'items_meta': [{'section': im['section'], 'label': im['label']} for im in items_meta],
    'visits': visits,
}
with open('dashboard_data.json', 'w') as f:
    json.dump(out, f, ensure_ascii=False)

print("visits:", len(visits))
print("items per visit:", len(visits[0]['items']))
print(json.dumps(visits[0], indent=2, ensure_ascii=False)[:1500])
