import pstats
p = pstats.Stats('stats')
p.sort_stats("tottime").print_stats()