import csv; from scipy.stats import pearsonr;
def read_rssi(path):
    data = []
    with open(path, 'r') as f:
        for row in csv.reader(f):
            if row: data.append(int(row[0]))
    return data
alice = read_rssi(r'data\alice\skenario1_mita_alice.csv')
bob = read_rssi(r'data\bob\skenario1_mita_bob.csv')
for i in range(1, 11):
    c = pearsonr(alice[:1000-i], bob[:1000-i])[0]
    print(f'Corr 1000-{i}: {c}')

