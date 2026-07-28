import pandas as pd
df = pd.read_pickle('df.pkl')
action_cols = [df.columns[14], df.columns[23], df.columns[46], df.columns[58]]
themes = {
 'Alumni Hand-holding gap': ['alumni handholding','alumni hand-holding','alumni hand holding'],
 'GROW Manual course pending': ['grow manual','grow process manual','skillfy','skillify'],
 'Weekly meeting not documented': ['weekly meeting','weekly team meeting'],
 'Mock Interview tracker gap': ['mock interview'],
 'Employer Connect / Top 20 gap': ['top 20','top20','employer connect'],
 'CCTV / scrap / infra': ['cctv','scrap'],
 'Batch start date not displayed': ['notice board','batch start date','batch-start date'],
 'Computer-generated marksheet': ['computer generated','computer-generated'],
 'CA diary not maintained': ['ca diary',"ca's diary",'ca dairy'],
 'AH/CA hiring delay or resignation': ['resign','hiring is delayed','hiring delayed','vacant','vacancy'],
}
counts={}
for name,kws in themes.items():
    centres=set()
    for idx,row in df.iterrows():
        for col in action_cols:
            val = row[col]
            if pd.isna(val): continue
            s = str(val).lower()
            if any(k in s for k in kws):
                centres.add(row['Centre Name'])
                break
    counts[name]=len(centres)
for k,v in sorted(counts.items(), key=lambda x:-x[1]):
    print(k, '->', v, 'centres')
