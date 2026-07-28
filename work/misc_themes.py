import pandas as pd
df = pd.read_pickle('df.pkl')
action_cols = [df.columns[14], df.columns[23], df.columns[46], df.columns[58]]
themes = {
 'computer generated marksheet': ['computer generated','computer-generated'],
 'alumni handholding gap': ['alumni handholding','alumni hand-holding','alumni hand holding'],
 'weekly meeting not documented': ['weekly meeting','weekly team meeting'],
 'GROW manual course pending': ['grow manual','grow process manual','skillfy','skillify'],
 'employer connect / top20 gap': ['top 20','top20','employer connect'],
 'mock interview tracker gap': ['mock interview'],
 'CCTV / scrap / infra': ['cctv','scrap'],
 'batch start date not displayed': ['notice board','batch start date','batch-start date'],
 'CA diary not maintained': ['ca diary',"ca's diary",'ca dairy'],
 'hostel food quality': ['hostel','food quality'],
}
counts={}
for name,kws in themes.items():
    c=0
    centres=set()
    for idx,row in df.iterrows():
        for col in action_cols:
            val = row[col]
            if pd.isna(val): continue
            s = str(val).lower()
            if any(k in s for k in kws):
                c+=1
                centres.add(row['Centre Name'])
                break
        else:
            continue
    counts[name]=(c,len(centres))
for k,v in sorted(counts.items(), key=lambda x:-x[1][0]):
    print(k, '-> mentions in', v[0],'rows across', v[1],'centres')
