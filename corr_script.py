import csv; from scipy.stats import pearsonr;
def read_rssi(path):
    data = []
    with open(path, 'r') as f:
        for row in csv.reader(f):
            if row: data.append(int(row[0]))
    return data
alice = read_rssi(r'data\alice\skenario1_mita_alice.csv')
bob = read_rssi(r'data\bob\skenario1_mita_bob.csv')
ln = min(len(alice), len(bob))
corr, _ = pearsonr(alice[:ln], bob[:ln])
print(corr)

