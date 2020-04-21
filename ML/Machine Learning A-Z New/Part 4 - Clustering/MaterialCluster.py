import pyodbc
import numpy as np
import matplotlib.pyplot as plt
import pandas as pd


conn = pyodbc.connect("Driver={SQL Server};"
                          "Server=10.118.23.78;"
                          "Database=D-DASH_Phase0_DQ;"
                          "UID=dsuser;"
                          "PWD=Password.1;"
                          # "Trusted_Connection=yes;"
                          )
cursor = conn.cursor()
Qry1111 = []

cursor.execute("select b.MATNR,mara.mtart,b.c,b.c1 from (select a.MATNR ,count(*) as c,count(*) as c1 from ( Select  mara.MATNR as MATNR,VBELN,VBAP.MATNR as VBAPMATNR from mara left join vbap on mara.MATNR=VBAP.MATNR  where VBELN is not NULL ) a group by a.matnr ) as b left join mara on b.matnr=mara.matnr UNION (Select  mara.MATNR as MATNR,mara.mtart,'0','0'  from mara left join vbap on mara.MATNR=VBAP.MATNR  where VBELN is NULL) order by 2 desc ")

names = list(map(lambda x: x[0], cursor.description))
Qry1111.append(names)
a = 0
for row in cursor.fetchall():
    l = []

    for i in range(0, len(row)):
        l.append(row[i])
    a=a+1
    Qry1111.append(l)

dataset=pd.DataFrame(Qry1111)

# K-Means Clustering

X = dataset.iloc[1:,[1,3]].values
# y = dataset.iloc[:, 3].values

from sklearn.preprocessing import LabelEncoder, OneHotEncoder
from sklearn.compose import ColumnTransformer
ct=ColumnTransformer([('encoder',OneHotEncoder(),[0])],remainder='passthrough')
X=np.array(ct.fit_transform(X),dtype=np.int)
'''

# Using the elbow method to find the optimal number of clusters
from sklearn.cluster import KMeans
wcss = []
for i in range(1, 11):
    kmeans = KMeans(n_clusters = i, init = 'k-means++', random_state = 0)
    kmeans.fit(X)
    wcss.append(kmeans.inertia_)
plt.plot(range(1, 11), wcss)
plt.title('The Elbow Method')
plt.xlabel('Number of clusters')
plt.ylabel('WCSS')
plt.show()

# Fitting K-Means to the dataset
kmeans = KMeans(n_clusters = 4, init = 'k-means++', random_state = 42)
y_kmeans = kmeans.fit_predict(X)

# Visualising the clusters
plt.scatter(X[y_kmeans == 0, 0], X[y_kmeans == 0, 0], s = 100, c = 'red', label = 'Cluster 1')
plt.scatter(X[y_kmeans == 1, 0], X[y_kmeans == 1, 0], s = 100, c = 'blue', label = 'Cluster 2')
plt.scatter(X[y_kmeans == 2, 0], X[y_kmeans == 2, 0], s = 100, c = 'green', label = 'Cluster 3')
plt.scatter(X[y_kmeans == 3, 0], X[y_kmeans == 3, 0], s = 100, c = 'cyan', label = 'Cluster 4')
#plt.scatter(X[y_kmeans == 4, 0], X[y_kmeans == 4, 1], s = 100, c = 'magenta', label = 'Cluster 5')
plt.scatter(kmeans.cluster_centers_[:, 0], kmeans.cluster_centers_[:, 1], s = 300, c = 'yellow', label = 'Centroids')
plt.title('Clusters of Materials')
plt.xlabel('Annual Income (k$)')
plt.ylabel('Spending Score (1-100)')
plt.legend()
plt.show()'''