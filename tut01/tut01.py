# tut1 solution
import csv
import pandas as pd
with open('octant_input.csv','r') as input_file:
    reader=csv.reader(input_file)
    with open('output.csv','w',newline='') as output_file:
        writer=csv.writer(output_file)
        for row in reader:
            writer.writerow(row)
data=pd.read_csv('output.csv')
u_avg=data['U'].mean()
v_avg=data['V'].mean()
w_avg=data['W'].mean()
print(u_avg)
print(v_avg)
print(w_avg)
# data['U_avg']=""
# data['V_avg']=""
# data['W_avg']=""
data.at[0,'U_avg']=u_avg
data.at[0,'V_avg']=v_avg
data.at[0,'W_avg']=w_avg
i=0
for ele in data['U']:
    data.at[i,"U'=U-Uavg"]=data.at[i,'U']-u_avg
    data.at[i,"U'=U-Vavg"]=data.at[i,'V']-v_avg
    data.at[i,"U'=U-Wavg"]=data.at[i,'W']-w_avg
    i=i+1



data.to_csv('output.csv',index=False)

