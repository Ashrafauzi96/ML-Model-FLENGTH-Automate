import cx_Oracle
import pandas as pd
import numpy as np
import itertools
import statsmodels.api as sm
from sklearn.model_selection import train_test_split
from sklearn.linear_model import LinearRegression
from sklearn.ensemble import RandomForestRegressor
from sklearn.metrics import mean_squared_error, r2_score
from sklearn.linear_model import Ridge
from sklearn.metrics import mean_squared_error

# data for train the ML model
df = pd.read_csv('C:/Users/fauziah/Desktop/export2.csv')

#x and y variable
X = df[['RECV_GREEN_LENGTH', 'PI_MEASLENGTH_MM', 'FIN_LENGTH_IN_MM', 'GREEN_DIAMETER']]
y = df['FIN_LENGTH_OUT_MM']

# Split the data into training and testing sets
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)

# Initialize the models
random_forest_model = RandomForestRegressor()
random_forest_model.fit(X_train, y_train)

# Get input data from database
import cx_Oracle
import sqlalchemy as sqla
import os
import pandas as pd
import numpy as np

conn = cx_Oracle.connect('xhq/xhq2mes@akuas952.office.graphiteelectrodes.net/mes301')
cur = conn.cursor()

sql_st = """
SELECT RECV_PIECE_ID, RECV_GREEN_LENGTH, PI_MEASLENGTH_MM, FIN_LENGTH_IN_MM, RECV_GREEN_DIAMETER_H AS GREEN_DIAMETER,FIN_LENGTH_OUT_MM,FIN_TIME_EXIT_S3,
(RECV_GREEN_LENGTH-PI_MEASLENGTH_MM) AS BAKESHRINKAGE,
(PI_MEASLENGTH_MM-FIN_LENGTH_IN_MM) AS GRAPHSHRINKAGE,
RECV_COKE,
RECV_AOS_CLUSTER,
BAK_FIRING_RECIPE,
GRA_SPIN_FORMULANUMBER,
RECV_NOMINAL_LENGTH,
RECV_NOMINAL_OD
FROM vw_mes_total
WHERE
(RECV_GREEN_LENGTH-PI_MEASLENGTH_MM)>0 AND (RECV_GREEN_LENGTH-PI_MEASLENGTH_MM)<80  AND 
(PI_MEASLENGTH_MM-FIN_LENGTH_IN_MM)>0 AND (PI_MEASLENGTH_MM-FIN_LENGTH_IN_MM)<80  AND 
RECV_GREEN_LENGTH>0 AND PI_MEASLENGTH_MM>0 AND FIN_LENGTH_IN_MM>0 AND RECV_GREEN_DIAMETER_H >0 AND
FIN_MACHINED_ENDDATE > '01-01-2020'
ORDER BY FIN_MACHINED_ENDDATE DESC  """

df_input = pd.read_sql_query(sql_st, con=conn)

#predict the ML Model with input data
df_input['prediction'] = random_forest_model.predict(df_input[['RECV_GREEN_LENGTH', 'PI_MEASLENGTH_MM', 'FIN_LENGTH_IN_MM', 'GREEN_DIAMETER']])

# export the result into excel
from sklearn.datasets import fetch_openml
import openpyxl

# Read dataset from OpenML
dataset = df_input
header = list(dataset.columns)
data = dataset.to_numpy().tolist()

# Create Excel workbook and write data into the default worksheet
wb = openpyxl.Workbook()
sheet = wb.create_sheet("Prediction3")  # or wb.active for default sheet
sheet.append(header)
for row in data:
    sheet.append(row)
# Save
wb.save("//akuas001/temp/Ashraf/Result_Prediction/Result.xlsx")



