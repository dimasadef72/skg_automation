import csv; from scipy.stats import pearsonr;
def read_rssi(path):
    data = []
    with open(path, 'r') as f:
        for row in csv.reader(f):
            if row: data.append(int(row[0]))
    return data
alice = read_rssi(r'data\alice\skenario1_mita_alice.csv')
bob = read_rssi(r'data\bob\skenario1_mita_bob.csv')
print(f'Full sync (1000 items): {pearsonr(alice, bob)[0]}')
print(f'Trimmed sync (998 items): {pearsonr(alice[:998], bob[:998])[0]}')

