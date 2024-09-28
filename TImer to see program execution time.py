import datetime
import time
import multiprocessing

def show_current_time():
    start_time = time.time()
    last_update_time = start_time
    while True:
        current_time = time.time()
        elapsed_time = current_time - start_time
        time_str = str(datetime.timedelta(seconds=elapsed_time))
        ms_str = str(int((elapsed_time % 1) * 1000)).zfill(3)
        print("Elapsed time: {}.{}s".format(time_str, ms_str), end="\r \n")
        time_to_sleep = max(0, last_update_time + 0.001 - current_time)
        time.sleep(time_to_sleep)
        last_update_time = current_time

def show_current_time2():
    start_time = time.time()
    last_update_time = start_time
    while True:
        current_time = time.time()
        elapsed_time = current_time - start_time
        time_str = str(datetime.timedelta(seconds=elapsed_time))
        ms_str = str(int((elapsed_time % 1) * 1000)).zfill(3)
        print("Elapsed time: {}.{}s".format(time_str, ms_str), end="\r")
        time_to_sleep = max(0, last_update_time + 0.001 - current_time)
        time.sleep(time_to_sleep)
        last_update_time = current_time        

# Example usage:
show_current_time()
show_current_time2()


p1 = multiprocessing.Process(target=show_current_time,args=[2])
p2  = multiprocessing.Process(target=show_current_time2,args=[2])

if __name__ == "__main__":
    p1.start()
    p2.start()
    p1.join()
    p2.join()

