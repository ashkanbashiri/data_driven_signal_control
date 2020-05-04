import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from sklearn import linear_model
from sklearn.preprocessing import PolynomialFeatures
from sklearn.preprocessing import StandardScaler
from sklearn.metrics import mean_squared_error, r2_score
from keras.layers import Dense, Activation,Dropout
from keras.models import Sequential
from keras.utils.vis_utils import plot_model
from keras.callbacks import TensorBoard


def build_model(df,batch_size,n_epochs):
    df.to_csv('./filtered_data_2.csv')
    test_df = pd.read_csv('./to_predict_2.csv')
    x_test = np.column_stack((test_df['flow1'].values.tolist(), test_df['flow1'].values.tolist(),
                         test_df['flow3'].values.tolist(), test_df['flow4'].values.tolist(),
                         test_df['lost_time'].values.tolist()))
    y = np.column_stack((df['green_time(1)'].values.tolist(), df['green_time(2)'].values.tolist(),
                         df['green_time(3)'].values.tolist(), df['green_time(4)'].values.tolist()))
    x = np.column_stack((df['flow(1)'].values.tolist(), df['flow(2)'].values.tolist(),
                         df['flow(3)'].values.tolist(), df['flow(4)'].values.tolist(),
                         df['losttime(s)'].values.tolist()))

    x_test_multiplier = x_test.max(axis=0)
    x_multiplier = x.max(axis=0)
    #x_test_multiplier = [2520, 2520, 2520, 2520, 36]
    print(x_multiplier)
    print(x_test_multiplier)
    x = x / x_multiplier
    x_test = x_test / x_multiplier
    y_multiplier = y.max(axis=0)
    print(y_multiplier)
    y = y / y_multiplier
    model = Sequential()
    model.add(Dense(units=512, activation='relu', input_dim=5))
    model.add(Dense(units=512, activation='relu'))
    model.add(Dense(units=512, activation='relu'))
    model.add(Dense(units=512, activation='relu'))
    model.add(Dense(units=512, activation='relu'))
    model.add(Dense(units=4, activation='linear'))
    model.compile(optimizer='adam', loss='mse')

    tensorboard = TensorBoard(log_dir='./logs', histogram_freq=0,
                              write_graph=True, write_images=False)
    history = model.fit(x, y, batch_size = batch_size, epochs = n_epochs, verbose=1, callbacks=[tensorboard])
    predictions = model.predict(x) * y_multiplier
    plt.plot(history.history['loss'])
    plt.title('model loss (mean squared error)')
    plt.ylabel('loss')
    plt.xlabel('epoch')
    plt.show()
    paper_predictions = model.predict(x_test) * y_multiplier
    out_df = pd.DataFrame(paper_predictions,columns=['green_1','green2','green3','green4'])
    out_df.to_csv('./paper_predictions_int2.csv')
    return predictions, model


def build_single_model(df,target,batch_size,n_epochs):
    y = df[target].values.tolist()
    x = np.column_stack((df['flow(1)'].values.tolist(), df['flow(2)'].values.tolist(),
                         df['flow(3)'].values.tolist(), df['flow(4)'].values.tolist()))
    x = x / np.amax(x)
    multiplier = np.amax(y)
    y = y / multiplier
    model = Sequential()
    model.add(Dense(256, activation='relu', input_dim = 4))

    for i in range(4):
        model.add(Dense(units=1024, activation='relu'))
        model.add(Dense(units=1024, activation='relu'))
    model.add(Dense(units=1, activation='linear'))
    model.compile(optimizer='adam', loss='mse')

    model.fit(x, y, batch_size = batch_size, epochs = n_epochs, verbose=2)
    my_ct_ann = model.predict(x) * multiplier
    return my_ct_ann