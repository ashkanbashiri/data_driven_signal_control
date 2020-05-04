from SALib.sample import saltelli
from SALib.analyze import sobol
from SALib.test_functions import Ishigami
import numpy as np

def analyse(model,n_samples):
    problem = {
        'num_vars': 5,
        'names': ['flow1', 'flow2', 'flow3', 'flow4', 'lost_time'],
        'bounds': [[0.0, 1.0],
                   [0.0, 1.0],
                   [0.0, 1.0],
                   [0.0, 1.0],
                   [0.0, 1.0],]
    }
    param_values = saltelli.sample(problem, n_samples)
    Y1 = np.zeros([param_values.shape[0]])
    Y2 = np.zeros([param_values.shape[0]])
    Y3 = np.zeros([param_values.shape[0]])
    Y4 = np.zeros([param_values.shape[0]])
    for i, X in enumerate(param_values):
        X = np.column_stack((X[0],X[1],X[2],X[3],X[4]))
        temp = model.predict(X)
        Y1[i] = temp[0][0]
        Y2[i] = temp[0][1]
        Y3[i] = temp[0][2]
        Y4[i] = temp[0][3]

    Si = sobol.analyze(problem, Y1)
    print('sensitivity analysis for phase split 1')
    print(Si['S1'])
    print(Si['ST'])
    print("x1-x2:", Si['S2'][0, 1])
    print("x1-x3:", Si['S2'][0, 2])
    print("x2-x3:", Si['S2'][1, 2])
    print("x3-x4:", Si['S2'][2, 3])

    Si = sobol.analyze(problem, Y2)
    print('sensitivity analysis for phase split 2')
    print(Si['S1'])
    print(Si['ST'])
    print("x1-x2:", Si['S2'][0, 1])
    print("x1-x3:", Si['S2'][0, 2])
    print("x2-x3:", Si['S2'][1, 2])
    print("x3-x4:", Si['S2'][2, 3])

    Si = sobol.analyze(problem, Y3)
    print('sensitivity analysis for phase split 3')
    print(Si['S1'])
    print(Si['ST'])
    print("x1-x2:", Si['S2'][0, 1])
    print("x1-x3:", Si['S2'][0, 2])
    print("x2-x3:", Si['S2'][1, 2])
    print("x3-x4:", Si['S2'][2, 3])

    Si = sobol.analyze(problem, Y4)
    print('sensitivity analysis for phase split 4')
    print(Si['S1'])
    print(Si['ST'])
    print("x1-x2:", Si['S2'][0, 1])
    print("x1-x3:", Si['S2'][0, 2])
    print("x2-x3:", Si['S2'][1, 2])
    print("x3-x4:", Si['S2'][2, 3])

    return True