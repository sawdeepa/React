e import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

dataset=pd.read_csv('Data.csv')
X=dataset.iloc[:,:-1].values
Y=dataset.iloc[:,-1].values#Y=dataset.iloc[:,3]
'''print(X)
print(Y)'''

from sklearn.impute import SimpleImputer
missingdata=SimpleImputer(missing_values=np.nan,strategy="mean",verbose=0)
missingdata=missingdata.fit(X[:,1:3])
X[:,1:3]=missingdata.transform(X[:,1:3])
'''print("After NAN is replaced")
print(X)
print(Y)'''

from sklearn.preprocessing import LabelEncoder,OneHotEncoder
from sklearn.compose import ColumnTransformer

ct=ColumnTransformer([('encoder',OneHotEncoder(),[0])],remainder='passthrough')
#X = np.array(ct.fit_transform(X), dtype=np.float)
X = np.array(ct.fit_transform(X), dtype=np.int)
Y=LabelEncoder().fit_transform(Y)
'''print("After OneHotEncoder")
print(X)
print(Y)'''

from sklearn.model_selection import train_test_split
X_train,X_test,Y_train,Y_test=train_test_split(X,Y,test_size=0.2,random_state=0)
'''print("After Split Train :")
print(X_train)
print(Y_train)
print("After Split Test :")
print(X_test)
print(Y_test)'''

from sklearn.preprocessing import StandardScaler
SC_X=StandardScaler()
X_train=SC_X.fit_transform(X_train)
X_test=SC_X.transform(X_train)
print(X_train)
print(X_test)
SC_Y=StandardScaler()
Y_train=SC_Y.fit_transform(Y_train.reshape(-1,1))
print(Y_train)
print(Y_test)

