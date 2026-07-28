import pandas as pd, re, json
pd.set_option('display.max_colwidth', None)
df = pd.read_excel('DRF_Unannounced Monthly Centre Visit by Second Level Managers (Responses).xlsx')

# Normalize manager names
def norm_mgr(x):
    x = str(x).strip()
    x = re.sub(r'\s+', ' ', x)
    mapping = {
        'phani':'Phani','s jakir':'S Jakir','raghavender rao':'Raghavender Rao','c.raghavender rao':'Raghavender Rao',
        'amit chakraborty':'Amit Chakraborty','bhargav challa':'Bhargav Challa','mahaboob basha j':'Mahaboob Basha',
        'mahaboob basha':'Mahaboob Basha','barna':'Barna','gagandeep singh sachar':'Gagandeep Singh Sachar',
        'gagandeeo singh sachar':'Gagandeep Singh Sachar','sanjay kumar tarai':'Sanjay Kumar Tarai',
        'c raghavenderrao':'Raghavender Rao'
    }
    key = x.lower().strip().rstrip('.')
    return mapping.get(key, x)

df['Manager'] = df["Second Level Manager’s Name"].apply(norm_mgr)
print(df['Manager'].value_counts())
print()
print("Total responses:", len(df))
print("Unique centres:", df['Centre Name'].nunique())
print("Unique managers:", df['Manager'].nunique())
df.to_pickle('/sessions/affectionate-busy-goodall/mnt/outputs/work/df.pkl')
