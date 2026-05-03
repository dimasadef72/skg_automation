import csv; from scipy.stats import pearsonr;
def read_rssi(path):
    data = []
    with open(path, 'r') as f:
        for row in csv.reader(f):
            if row: data.append(int(row[0]))
    return data
alice = read_rssi(r'data\alice\skenario1_mita_alice.csv')
bob = read_rssi(r'data\bob\skenario1_mita_bob.csv')
print(f'Corr All: {pearsonr(alice[:1000], bob[:1000])[0]}')
print(f'Corr 999: {pearsonr(alice[:999], bob[:999])[0]}')
print(f'Corr Skip 1: {pearsonr(alice[1:1000], bob[1:1000])[0]}')

