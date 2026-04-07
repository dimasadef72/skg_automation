import time
import numpy as np

# =====================================================================
# KALMAN PARAMETERS
# =====================================================================
KALMAN_A = 1
KALMAN_H = 1
KALMAN_XAPOSTERIORI_0 = -5
KALMAN_PAPOSTERIORI_0 = 1

def process_kalman(raw_data, q, r, bb, benchmark_iterations=10):
    """
    Melakukan proses sinyal dengan Kalman Filter. 
    Mengembalikan data filter rata rata, rata rata KGR, dan rata rata Execution Time dari 10 kali (parameter iteration) percobaan agar benchmark valid.
    """
    total_data = len(raw_data)
    if total_data < bb: return [], 0, 0.0
        
    aa = total_data // bb
    data_cut = raw_data[:aa * bb]
    signal = np.array(data_cut).reshape(aa, bb).T
    
    def run_kalman(sig):
        xaposteriori = []
        paposteriori = []
        row1, row2, row3, row4, row5, row6 = [], [], [], [], [], []
        
        for m in range(aa):
            row1.append(KALMAN_A * KALMAN_XAPOSTERIORI_0)
            row2.append(sig[0][m] - KALMAN_H * row1[m])
            row3.append(KALMAN_A * KALMAN_A * KALMAN_PAPOSTERIORI_0 + q)
            gain = row3[m] / (row3[m] + r)
            row4.append(gain)
            row5.append(row3[m] * (1 - gain))
            row6.append(row1[m] + gain * row2[m])
            
        xaposteriori.append(row6)
        paposteriori.append(row5)
        
        for j in range(1, bb):
            r1, r2, r3, r4, r5, r6 = [], [], [], [], [], []
            for n in range(aa):
                r1.append(xaposteriori[j-1][n])
                r2.append(sig[j][n] - KALMAN_H * r1[n])
                r3.append(KALMAN_A * KALMAN_A * paposteriori[j-1][n] + q)
                gain = r3[n] / (r3[n] + r)
                r4.append(gain)
                r5.append(r3[n] * (1 - gain))
                r6.append(r1[n] + gain * r2[n])
            xaposteriori.append(r6)
            paposteriori.append(r5)
        return xaposteriori

    times = []
    kgrs = []
    hasil_array = []
    gain_count = aa * bb
    
    for _ in range(benchmark_iterations):
        start = time.perf_counter()
        result = run_kalman(signal)
        end = time.perf_counter()
        elapsed = end - start
        
        if elapsed == 0: elapsed = 1e-9  # Avoid exact zero div
        kgr = gain_count * 32 / elapsed
        
        times.append(elapsed)
        kgrs.append(kgr)
        hasil_array = result
        
    avg_times = np.mean(times)
    avg_kgrs = np.mean(kgrs)
    
    # Flatten dari array matrix memanjang dan casting ke Integer seperti script asli
    hasil_flat = np.array(hasil_array).T.reshape(-1)
    hasil_int = [int(val) for val in hasil_flat]
    return hasil_int, avg_kgrs, avg_times
