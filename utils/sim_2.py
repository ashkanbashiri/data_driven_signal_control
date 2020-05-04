import win32com.client as com
import os
import sys

global Vissim
import time
import math



def generate_phase_times(ct, sf, flows, lt, n_phase, scheme='relative'):
    phase_times = [0] * 3
    if scheme == 'relative':
        for i in range(0,3):
            phase_times[i] = (flows[i]) * (ct - lt)

    elif scheme == 'fair':
        for i in range(0, 3):
            phase_times[i] = (0.33) * (ct - lt)
    return phase_times


def run_simulation(ct, phase_times, lost_time, flows, first_time, results_file, close_vissim=False, reset_vissim=False):
    import win32com.client as com
    global Vissim
    sim_period = 30 * 60  # 60 minutes
    num_cycles = int(sim_period / ct)
    num_lanes = 4
    sat_rate = 1900
    saturation_flow_rate = num_lanes * sat_rate  # veh/hour
    flows = [x * saturation_flow_rate for x in flows]
    # print flows
    vi_numbers = [1, 2, 3, 4]
    if first_time == 1 or reset_vissim is True:
        Vissim = com.gencache.EnsureDispatch("Vissim.Vissim")
        Path_of_COM_Basic_Commands_network = os.getcwd()
        Filename = os.path.join(Path_of_COM_Basic_Commands_network, './vissim/large_int.inpx')
        flag_read_additionally = False
        Vissim.LoadNet(Filename, flag_read_additionally)

    Vissim.Simulation.SetAttValue('SimBreakAt', 1)
    Vissim.Simulation.SetAttValue('SimPeriod', sim_period)
    #Vissim.Simulation.SetAttValue('SimSpeed', 1)
    Vissim.Simulation.SetAttValue('UseMaxSimSpeed', True)
    Vissim.Graphics.CurrentNetworkWindow.SetAttValue("QuickMode", 1)
    Vissim.SuspendUpdateGUI()
    #time.sleep(1)
    Vissim.Simulation.RunContinuous()

    # start the simulation until SimBreakAt (1s)
    '''
    #phase1 = 1 - 1 1 - 3 2 - 1 2 - 2
    #phase2 = 3 - 1 3 - 2 4 - 1 4 - 2
    #phase3 =  1 - 2 2 - 3
    #phase4 = 3 - 3 4 - 3
    static_flows = [.3*flows[0],.2*flows[0],.3*flows[0],.2*flows[0],.6*flows[1],.6*flows[1],.6*flows[1],
                    .6*flows[1],.5*flows[2],.5*flows[2],.5*flows[3],.5*flows[3],]
    #Setting Relative flow for simulation#2
    decision_numbers = [1,1,2,2,3,3,4,4,1,2,3,4]
    routing_numbers = [1,3,1,2,1,2,1,2,2,3,3,3]
    counter = 0

    for i in xrange(len(decision_numbers)):
        Vissim.Net.VehicleRoutingDecisionsStatic.ItemByKey(decision_numbers[counter]).VehRoutSta.ItemByKey(routing_numbers[counter]).SetAttValue(
        'RelFlow(1)', static_flows[counter])
    '''
    rel_flows = [0, 0, 0]
    for i in range(0, 3):
        rel_flows[i] = flows[i] / sum(flows)
    static_flows = [0.2,0.8, 0.8*rel_flows[0]/(rel_flows[0]+ rel_flows[2]),0.2*rel_flows[0]/(rel_flows[0]+ rel_flows[2]),
                    rel_flows[2]/(rel_flows[0]+ rel_flows[2]),0.8,0.2,
                    0.2 * rel_flows[0] / (rel_flows[0] + rel_flows[2]), 0.8 * rel_flows[0] / (rel_flows[0] + rel_flows[2]),
                    rel_flows[2] / (rel_flows[0] + rel_flows[2])
                    ]
    counter = 0
    decision_numbers = [1, 1, 2, 2, 2, 3, 3, 4, 4, 4]
    routing_numbers = [1, 2, 1, 2, 3, 1, 2, 1, 2, 3]
    for i in range(len(decision_numbers)):
        Vissim.Net.VehicleRoutingDecisionsStatic.ItemByKey(decision_numbers[i]).VehRoutSta.ItemByKey(
            routing_numbers[i]).SetAttValue(
            'RelFlow(1)', static_flows[i])

    leg_flows = [0.5*flows[1],0.5*flows[0] + 0.5*flows[2] ,0.5*flows[1] ,0.5*flows[0] + 0.5*flows[2]]
    for vi_number in vi_numbers:
        veh_input = Vissim.Net.VehicleInputs.ItemByKey(vi_number)
        veh_input.SetAttValue('Volume(1)', leg_flows[vi_number - 1])
    signal_controller = Vissim.Net.SignalControllers.ItemByKey(1)
    sgs = [1,2,3]
    green_times = phase_times
    break_time = 1
    green = "GREEN"
    yellow = "AMBER"  # possible values 'GREEN', 'RED', 'AMBER', 'REDAMBER' and more, see COM Help: SignalizationState Enumeration
    red = "RED"
    yellow_time = lost_time / 3
    for cycle in range(num_cycles):
        cntr = 0
        for sg in sgs:
            green_time = math.ceil(green_times[cntr])
            cntr += 1
            break_time += green_time
            if break_time < sim_period:
                signal_group = signal_controller.SGs.ItemByKey(sg)
                signal_group.SetAttValue("SigState", green)
                for other in [other for other in sgs if other != sg]:
                    signal_group = signal_controller.SGs.ItemByKey(other)
                    signal_group.SetAttValue("SigState", red)
                Vissim.Simulation.SetAttValue('SimBreakAt', break_time)
                Vissim.Simulation.RunContinuous()
                break_time += yellow_time
                if break_time < sim_period:
                    signal_group = signal_controller.SGs.ItemByKey(sg)
                    signal_group.SetAttValue("SigState", red)
                    Vissim.Simulation.SetAttValue('SimBreakAt', break_time)
                    Vissim.Simulation.RunContinuous()
    avg_delay = Vissim.Net.Nodes.ItemByKey(1).TotRes.AttValue('VehDelay (Current,Avg,All)')
    stdev_delay = Vissim.Net.Nodes.ItemByKey(1).TotRes.AttValue('VehDelay (Current,StdDev,All)')
    q_len = Vissim.Net.Nodes.ItemByKey(1).TotRes.AttValue('QLen (Current,Avg)')
    stops = Vissim.Net.Nodes.ItemByKey(1).TotRes.AttValue('Stops (Current,Avg,All)')
    stop_delay = Vissim.Net.Nodes.ItemByKey(1).TotRes.AttValue('StopDelay (Current,Avg,All)')
    emmissions_co = Vissim.Net.Nodes.ItemByKey(1).TotRes.AttValue('EmissionsCO (Current,Avg)')
    fuel_consumption = Vissim.Net.Nodes.ItemByKey(1).TotRes.AttValue('FuelConsumption (Current,Avg)')

    csv_filename = results_file
    columnTitleRow = "cycletime(s), totalflow(Veh/h),flow(1),flow(2),flow(3),losttime(s)," \
                     "green_time(1),green_time(2),green_time(3),avgdelay,stopsPerVehicle,stopDelay," \
                     "emissionsco,fuelconsumption,stdev_delay,avg_qlen\n"
    row_data = "{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{}\n".format(ct, sum(flows), flows[0], flows[1],
                                                                                flows[2], lost_time,
                                                                                green_times[0], green_times[1],
                                                                                green_times[2],
                                                                                avg_delay, stops, stop_delay,
                                                                                emmissions_co, fuel_consumption,
                                                                                stdev_delay, q_len)

    if first_time == 1:
        with open(csv_filename, 'w') as csv:
            csv.write(columnTitleRow)
    with open(csv_filename, 'a') as csv:
        csv.write(row_data)
    Vissim.Simulation.Stop()
    if close_vissim is True:
        Vissim = None


def printProgressBar(iteration, total, prefix='', suffix='', decimals=1, length=100, fill='#'):
    """
    Call in a loop to create terminal progress bar
    @params:
        iteration   - Required  : current iteration (Int)
        total       - Required  : total iterations (Int)
        prefix      - Optional  : prefix string (Str)
        suffix      - Optional  : suffix string (Str)
        decimals    - Optional  : positive number of decimals in percent complete (Int)
        length      - Optional  : character length of bar (Int)
        fill        - Optional  : bar fill character (Str)
    """
    percent = ("{0:." + str(decimals) + "f}").format(100 * (iteration / float(total)))
    filledLength = int(length * iteration // total)
    bar = fill * filledLength + '-' * (length - filledLength)
    print('\r%s |%s| %s%% %s\r' % (prefix, bar, percent, suffix))
    sys.stdout.flush()

    # Print New Line on Complete
    if iteration == total:
        print()