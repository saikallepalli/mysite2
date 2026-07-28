import pandas as pd, re, json
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

items_meta = []
for sec, idxs in sections.items():
    for i in idxs:
        col = df.columns[i]
        items_meta.append({'section': sec, 'label': clean_label(col), 'idx': i})

# ---- Long format Data table ----
long_rows = []
visit_wide_rows = []
qcols = [df.columns[59], df.columns[60], df.columns[61]]

for vi, row in df.iterrows():
    manager = row['Manager']; centre = row['Centre Name']; program = row['Program Name']
    date = row['Visit Date']
    date_s = pd.Timestamp(date).strftime('%Y-%m-%d') if pd.notna(date) else ''
    sec_counts = {s: {'AW':0,'NI':0,'NA':0} for s in sections}
    for im in items_meta:
        val = row[df.columns[im['idx']]]
        status = str(val) if pd.notna(val) else 'Not Applicable'
        long_rows.append({'Manager': manager, 'Centre': centre, 'Program': program, 'Date': date_s,
                           'Section': im['section'], 'Item': im['label'], 'Status': status})
        if status == 'All is Well': sec_counts[im['section']]['AW'] += 1
        elif status == 'Need Improvement': sec_counts[im['section']]['NI'] += 1
        else: sec_counts[im['section']]['NA'] += 1
    wide = {
        'Manager': manager, 'Centre': centre, 'Program': program, 'Date': date_s,
        'DataCheck_AW': sec_counts['Data Check']['AW'], 'DataCheck_NI': sec_counts['Data Check']['NI'], 'DataCheck_NA': sec_counts['Data Check']['NA'],
        'InfraCheck_AW': sec_counts['Infra Check']['AW'], 'InfraCheck_NI': sec_counts['Infra Check']['NI'], 'InfraCheck_NA': sec_counts['Infra Check']['NA'],
        'ProcessCheck_AW': sec_counts['Process Check']['AW'], 'ProcessCheck_NI': sec_counts['Process Check']['NI'], 'ProcessCheck_NA': sec_counts['Process Check']['NA'],
        'AspirantInt_AW': sec_counts['Aspirant Interaction']['AW'], 'AspirantInt_NI': sec_counts['Aspirant Interaction']['NI'], 'AspirantInt_NA': sec_counts['Aspirant Interaction']['NA'],
        'Q_Bottleneck': (row[qcols[0]] if pd.notna(row[qcols[0]]) else ''),
        'Q_PlacementAction': (row[qcols[1]] if pd.notna(row[qcols[1]]) else ''),
        'Q_BatchOnTimeAction': (row[qcols[2]] if pd.notna(row[qcols[2]]) else ''),
    }
    visit_wide_rows.append(wide)

data_long = pd.DataFrame(long_rows)
visits_wide = pd.DataFrame(visit_wide_rows)

unique_items = pd.DataFrame(items_meta)[['section','label']].drop_duplicates().rename(columns={'section':'Section','label':'Item'})

data_long.to_csv('data_long.csv', index=False)
visits_wide.to_csv('visits_wide.csv', index=False)
unique_items.to_csv('unique_items.csv', index=False)

print('data_long', data_long.shape)
print('visits_wide', visits_wide.shape)
print('unique_items', unique_items.shape)
print(sorted(data_long['Manager'].unique()))
print(len(sorted(data_long['Centre'].unique())))
