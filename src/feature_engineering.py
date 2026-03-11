#1.Load the train and test data
#2.scale the training data
#3.save the scaled data to the processed folder

from datapreprocessing import load_and_Split_data
from sklearn.preprocessing import StandardScaler
import pandas as pd
import pickle

x_train, x_test, y_train, y_test = load_and_Split_data()

scaler = StandardScaler()
x_train_scaled = scaler.fit_transform(x_train)
x_test_scaled = scaler.transform(x_test)

pd.DataFrame(x_train_scaled).to_csv("../data/processed/X_train.csv", index=False)
pd.DataFrame(x_test_scaled).to_csv("../data/processed/X_test.csv", index=False)
pd.DataFrame(y_train).to_csv("../data/processed/y_train.csv", index=False)
pd.DataFrame(y_test).to_csv("../data/processed/y_test.csv", index=False)

with open("../artifacts/scaler.pkl","wb") as f:
    pickle.dump(scaler,f)
print("successfully saved your scaler.pkl file")