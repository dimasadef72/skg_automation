import time
import numpy as np

def gray_code(n):
    """Fungsi pembangun kode gray untuk n byte konstan"""
    if n <= 0: return ['0']
    codes = ['0', '1']
    for bits in range(2, n+1):
        mirror = codes[::-1]
        codes = ['0' + c for c in codes] + ['1' + c for c in mirror]
    return codes

def process_kuantisasi(data, num_bits=3, benchmark_iterations=10, reference_min=None, reference_max=None):
    """
    Kuantisasi nilai array menjadi bitstream String dengan Gray Code.
    Mengembalikan Array Bitstream Final, rata rata KGR, dan rata rata Waktu (diuji iterasi 10x seperti aslinya).
    """
    def run_kuantisasi(data_arr):
        t0 = time.perf_counter()
        data_arr = np.asarray(data_arr).astype(float)
        n = len(data_arr)
        if n == 0: return "", 0, 0.0, 0.0
        
        levels = 2 ** num_bits
        if reference_min is None or reference_max is None:
            d_min = np.min(data_arr)
            d_max = np.max(data_arr)
        else:
            d_min = float(reference_min)
            d_max = float(reference_max)
        d_range = d_max - d_min
        
        if d_range == 0:
            indices = np.zeros(n, dtype=int)
        else:
            step = d_range / levels
            indices = np.floor((data_arr - d_min) / step).astype(int)
            indices[indices >= levels] = levels - 1
            
        gray_map = gray_code(num_bits)
        bitstream = "".join([gray_map[i] for i in indices])
        total_bits = len(bitstream)
        
        t1 = time.perf_counter()
        elapsed = t1 - t0
        if elapsed == 0: elapsed = 1e-9
        kgr = total_bits / elapsed
        return bitstream, total_bits, kgr, elapsed

    times, kgrs = [], []
    final_bitstream = ""
    
    for _ in range(benchmark_iterations):
        bstream, tbits, vkgr, vtime = run_kuantisasi(data)
        times.append(vtime)
        kgrs.append(vkgr)
        final_bitstream = bstream
        
    return final_bitstream, np.mean(kgrs), np.mean(times)
