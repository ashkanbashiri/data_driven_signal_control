import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from sklearn import linear_model
from sklearn.preprocessing import PolynomialFeatures

from sklearn.metrics import mean_squared_error, r2_score
from keras.layers import Dense, Activation
from keras.models import Sequential


def webster(flowratio,losttime):
    cycle_time = ((1.5*losttime)+5)/(1-flowratio)
    return cycle_time
def webster2(pair):#pair is a numpy 2d array, column 0 is flowratio, column two is losttime
    cycle_times = ((1.5*pair[:,1])+5)/(1-pair[:,0])
    return cycle_times

test_lost_time = 8

pd.set_option('display.max_rows', 500)
pd.set_option('display.max_columns', 500)
pd.set_option('display.width', 1000)
df = pd.read_csv('./Final/res/simulation_results_v1.csv', sep=',')
#optimal_cycles = df.loc[df.groupby(['totalflowratio','losttime','flow1','flow2','flow3','flow4'])["avgdelay"].idxmin()]
mean_col = df.groupby(['cycletime','totalflowratio','losttime'])['avgdelay'].mean()
df = df.set_index(['cycletime','totalflowratio','losttime']) # make the same index here
df['mean_col'] = mean_col
df = df.reset_index()
optimal_cycles = df.loc[df.groupby(['totalflowratio','losttime','flowratio1','flowratio2','flowratio3','flowratio4'])["mean_col"].idxmin()].drop_duplicates(subset=['cycletime','totalflowratio','losttime'],keep='first')
y = optimal_cycles['cycletime'].values.tolist()

x = np.column_stack((optimal_cycles['totalflowratio'].values.tolist(), optimal_cycles['losttime'].values.tolist()))
x_and_y = np.column_stack((x,y))
#x_and_y = np.column_stack((optimal_cycles['totalflowratio'].values.tolist(), optimal_cycles['losttime'].values.tolist(),optimal_cycles['cycletime'].values.tolist()))
'''
df2 = pd.read_csv('./Final/res/simulation_results_v1.csv', sep=',')
#optimal_cycles = df.loc[df.groupby(['totalflowratio','losttime','flow1','flow2','flow3','flow4'])["avgdelay"].idxmin()]
mean_col2 = df2.groupby(['cycletime','totalflowratio','losttime'])['avgdelay'].mean()
df2 = df2.set_index(['cycletime','totalflowratio','losttime']) # make the same index here
df2['mean_col'] = mean_col2
df2 = df2.reset_index()
optimal_cycles2 = df2.loc[df2.groupby(['totalflowratio','losttime','flow1','flow2','flow3','flow4'])["mean_col"].idxmin()].drop_duplicates(subset=['cycletime','totalflowratio','losttime'],keep='first')
y2 = optimal_cycles2['cycletime'].values.tolist()

x2 = np.column_stack((optimal_cycles2['totalflowratio'].values.tolist(), optimal_cycles2['losttime'].values.tolist()))
'''

for_plot = np.zeros(shape=(14,2))
cntr = 0
for m in range(len(x_and_y)):

    if np.allclose(x_and_y[m,[1]],[test_lost_time]):
        for_plot[cntr,:] = x_and_y[m,[0,2]]
        cntr+=1
#print(for_plot)


'''
poly = PolynomialFeatures(degree=2)
x_ = poly.fit_transform(x)
clf = linear_model.LinearRegression()
clf.fit(x_, y)
poly_predictions = clf.predict(x_)
print('Coefficients: \n', clf.coef_)

print("Mean squared error: %.2f"
      % mean_squared_error(y, poly_predictions))
print('R2 Score: %.2f' % r2_score(y, poly_predictions))

'''

x_test = np.zeros(shape=(13,2))
cntrr =0
for i in range(13):
    j = 30 + cntrr*5
    cntrr +=1
    x_test[i,:] = [(j)/100,test_lost_time]
webster_ct = webster2(x_test)
print(webster_ct)
#x_test_ = poly.fit_transform(x_test)
#my_ct_polynomial = clf.predict(x_test_)


# Initialising the ANN
model = Sequential()
model.add(Dense(512, activation='relu', input_dim = 2))
model.add(Dense(units=256, activation='relu'))
model.add(Dense(units=256, activation='relu'))
model.add(Dense(units=256, activation='relu'))
model.add(Dense(units=256, activation='relu'))
model.add(Dense(units=256, activation='relu'))
model.add(Dense(units = 1))
model.compile(optimizer = 'adam',loss = 'mean_squared_error')
x_normed = x# / x.max(axis=0)

model.fit(x_normed, y, batch_size = 150, epochs = 5000,verbose=0)
x_test_normed = x_test #/ x_test.max(axis=0)
my_ct_ann = model.predict(x_test_normed)
print(my_ct_ann)
##############################################################################
# Initialising the ANN
'''
model = Sequential()
model.add(Dense(16, activation='relu', input_dim = 2))
model.add(Dense(units=128, activation='relu'))
model.add(Dense(units=512, activation='relu'))
model.add(Dense(units = 1))
model.compile(optimizer = 'adam',loss = 'MSE')
model.fit(x2, y2, batch_size = 100, epochs = 2000,verbose=2)
my_ct_ann2 = model.predict(x_test)
'''
###################################################
#print(my_ct_ann)

fig, ax = plt.subplots()
webs = ax.plot(x_test[:,0], webster_ct, color='blue', linewidth=2,label ='Webster')

ax.set(xlabel='Total Flow Ratio', ylabel='Optimal Cycle Length (s)', title='Optimal Cycle Lengths for Lost Time = {}s'.format(test_lost_time))
mine_ann = ax.plot(x_test[:,0], my_ct_ann, color='green', linewidth=2, label="Neural Regression- 2nd Intersection")
#simulation = ax.plot(for_plot[:,0],for_plot[:,1], color='black', linewidth=2,label ='Simulation')

#mine_ann2 = ax.plot(x_test[:,0], my_ct_ann2, color='red', linewidth=2, label="Neural Regression - 2nd Intersection")

ax.legend()
ax.grid()


plt.show()