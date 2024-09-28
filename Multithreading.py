import time
import multiprocessing

def x(s):
    print(f"Sleeping for {s} seconds")
    time.sleep(s)
    print(f"Done sleeping {s} seconds")

p1 = multiprocessing.Process(target=x,args=[2])
p2  = multiprocessing.Process(target=x,args=[2])

if __name__ == "__main__":
    p1.start()
    p2.start()
    p1.join()
    p2.join()

finish = time.perf_counter()
print("finished running after seconds: ", finish)    