from sklearn.model_selection import train_test_split
import pandas as pd
import numpy as np
from keras.models import Model
from keras.utils import plot_model
from keras.layers import Dense, Flatten, Conv1D, MaxPooling1D, Input


def load_total_data(path, num):
    df = pd.read_pickle(path)
    x = np.expand_dims(df.values[:, 0:-1].astype(float), axis=2)  # Adding a one-dimensional axis
    y = df.values[:, -1] / num
    # Divide training set, test set
    x_train, x_test, y_train, y_test = train_test_split(x, y, test_size=0.05, shuffle=True)
    print("Loading of data complete!")
    return x_train, x_test, y_train, y_test

def build_CNN_model(model_structure):
    input1 = Input(shape=(187, 1))
    conv_layer1_1 = Conv1D(16, 3, strides=1, activation='relu')(input1)
    conv_layer1_2 = Conv1D(16, 3, strides=1, activation='relu')(conv_layer1_1)
    conv_layer1_3 = Conv1D(16, 3, strides=1, activation='relu')(conv_layer1_2)
    max_layer1 = MaxPooling1D(3)(conv_layer1_3)
    conv_layer2_1 = Conv1D(32, 3, strides=1, activation='relu')(max_layer1)
    conv_layer2_2 = Conv1D(32, 3, strides=1, activation='relu')(conv_layer2_1)
    conv_layer2_3 = Conv1D(32, 3, strides=1, activation='relu')(conv_layer2_2)
    max_layer2 = MaxPooling1D(3)(conv_layer2_3)
    flatten = Flatten()(max_layer2)
    f1 = Dense(1, activation='linear', name='prediction_one')(flatten)
    model = Model(outputs=f1, inputs=input1)
    model.summary()
    plot_model(model, to_file=model_structure, show_shapes=True)  # Printed model structure
    return model

def run(model_structure):
    # data_nh3_pata = 'Processed_Sin_NH3_data/NH3_average_removebd-Sinr-single-Sin-alldata.pkl'
    data_nh3_pata = 'Processed_NH3_data/nh3-spectrum.pkl'

    x_train, x_test, y_train, y_test = load_total_data(data_nh3_pata, 10000)
    model = build_CNN_model(model_structure)
    return model, x_train, x_test, y_train, y_test, 10000
