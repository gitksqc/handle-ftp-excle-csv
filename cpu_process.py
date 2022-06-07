import psutil

def fib(n):
    if n <= 2:
        return n
    else:
        return fib(n-1) + fib(n-2)

p = psutil.Process()
p.cpu_affinity([15,16])
fib(8)