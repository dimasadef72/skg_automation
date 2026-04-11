import numpy as np
import csv

# Generate 128x128 universal hash table (random but fixed seed)
np.random.seed(12345)  # IMPORTANT: keep seed fixed so Alice/Bob use same hashtable

table = np.random.randint(0, 2, size=(128, 128))

with open("Hashtable128.csv", "w", newline="") as f:
    w = csv.writer(f)
    for row in table:
        w.writerow(row)

print("✅ Hashtable128.csv telah dibuat! (128x128)")
