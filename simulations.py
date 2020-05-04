from utils.sim import generate_phase_times, run_simulation, printProgressBar
import sys
import time
import datetime
import multiprocessing
import win32com.client as com
import os


def main():
    global start_time
    global Vissim
    start_time = time.time()
    fair_scheme = 'fair'
    relative_scheme = 'relative'
    schemes = [fair_scheme,relative_scheme]
    n_phase = 4
    c_times = list(range(50,180,10))#CYCLE TIMES = [50,55,...,150] seconds
    sum_flows = list(range(30,101,10)) # TOTAL_FLOWS = [35,...,50,51,....,99,100]% of Saturation Flow Rate
    sum_flows = [float(x)/100 for x in sum_flows]
    lost_times = list(range(4,12,4))
    flow_ratios = [ [0.25, .25, .25, .25], [0.1, 0.4, 0.1, 0.4], [0.3, 0.3, 0.2, 0.2], [0.1, 0.1, 0.2, 0.6]]
    first_time = 1
    num_simulations = len(sum_flows)*len(lost_times)*len(c_times)*len(flow_ratios)*2
    print("Starting to run {} simulations".format(num_simulations))
    sim_number = 0
    total_flows = [1800, 3600,]
    last_checkpoint = 58
    close_vis = False
    restart_vissim = False
    for sum_flow in sum_flows:
        for lost_time in lost_times:
            for ct in c_times:
                for i in range(len(flow_ratios)):
                    ratios = [x * sum_flow for x in flow_ratios[i][:]]
                    for scheme_number in range(2):
                        _progress(sim_number, num_simulations)
                        phase_times = generate_phase_times(ct,sum_flow,flow_ratios[i][:],lost_time,n_phase,scheme=schemes[scheme_number])
                        if sim_number>last_checkpoint:
                            if close_vis is True:
                                restart_vissim = True
                            else:
                                restart_vissim = False
                            if sim_number%102 == 0:
                                close_vis = True
                            else:
                                close_vis = False
                            run_simulation(ct,phase_times,lost_time,ratios,first_time, 'results_nov18_final_2.csv',
                                           close_vissim=close_vis, reset_vissim=restart_vissim)
                            first_time = 0
                        sim_number += 1
    sys.stdout.write('\rAll simulations Completed Sucessfully!')


def _progress(count, total_size):
    global start_time
    elapsed_time = int(time.time() - start_time)
    sys.stdout.write('\r%.2f%% Completed, ' % (float(count) / float(total_size) * 100.0) +
                     '\tElapsed Time: {}'.format(str(datetime.timedelta(seconds=elapsed_time))))
    sys.stdout.flush()


if __name__ == "__main__":
    main()
