import pandas as pd, numpy as np
df = pd.read_pickle('df.pkl')
action_cols = [df.columns[14], df.columns[23], df.columns[46], df.columns[58]]
def wc(x):
    if pd.isna(x): return 0
    return len(str(x).split())
df['action_words'] = df[action_cols].applymap(wc).sum(axis=1)
df['action_filled'] = df[action_cols].notna().sum(axis=1)

g = df.groupby('Manager').agg(visits=('Manager','size'), avg_words=('action_words','mean'), avg_fields_filled=('action_filled','mean')).round(1).sort_values('avg_words', ascending=False)
print(g)
