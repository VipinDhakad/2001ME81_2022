# tut1 solution
import csv
import pandas as pd
with open('octant_input.csv','r') as input_file:
    reader=csv.reader(input_file)
    with open('octant_output.csv','w',newline='') as output_file:
        writer=csv.writer(output_file)
        for row in reader:
            writer.writerow(row)
data=pd.read_csv('octant_output.csv')
u_avg=data['U'].mean()
v_avg=data['V'].mean()
w_avg=data['W'].mean()
# print(u_avg)
# print(v_avg)
# print(w_avg)
# data['U_avg']=""
# data['V_avg']=""
# data['W_avg']=""
data.at[0,'U_avg']=u_avg
data.at[0,'V_avg']=v_avg
data.at[0,'W_avg']=w_avg
i=0
cnt_p1=0
cnt_n1=0
cnt_p2=0
cnt_n2=0
cnt_p3=0
cnt_n3=0
cnt_p4=0
cnt_n4=0

for ele in data['U']:
    x=data.at[i,"U'=U-Uavg"]=data.at[i,'U']-u_avg
    y=data.at[i,"V'=V-Vavg"]=data.at[i,'V']-v_avg
    z=data.at[i,"W'=W-Wavg"]=data.at[i,'W']-w_avg
    if x>0:
        if y>0:
            if z>0:
                data.at[i,'Octant']=1
                cnt_p1=cnt_p1+1
            else:
                data.at[i,'Octant']=-1
                cnt_n1=cnt_n1+1
        else:
            if z>0:
                data.at[i,'Octant']=4
                cnt_p4=cnt_p4+1
            else:
                data.at[i,'Octant']=-4
                cnt_n4=cnt_n4+1
    else:
        if y>0:
            if z>0:
                data.at[i,'Octant']=2
                cnt_p2=cnt_p2+1
            else:
                data.at[i,'Octant']=-2
                cnt_n2=cnt_n2+1
        else:
            if z>0:
                data.at[i,'Octant']=3
                cnt_p3=cnt_p3+1
            else:
                data.at[i,'Octant']=-3
                cnt_n3=cnt_n3+1
    i=i+1
data.at[0,'Octant ID']='Overall Count'
data.at[0,'1']=cnt_p1
data.at[0,'-1']=cnt_n1
data.at[0,'2']=cnt_p2
data.at[0,'-2']=cnt_n2
data.at[0,'3']=cnt_p3
data.at[0,'-3']=cnt_n3
data.at[0,'4']=cnt_p4
data.at[0,'-4']=cnt_n4

mod=5000
i=0
prev=0
iter=1
cnt_p1=0
cnt_n1=0
cnt_p2=0
cnt_n2=0
cnt_p3=0
cnt_n3=0
cnt_p4=0
cnt_n4=0
data.at[1,'Octant ID']=mod
# print(len(data))
while i<len(data):
    flag=False
    for j in range(prev,mod*iter):
        if i<len(data):
            if data.at[i,'Octant']==1:
                cnt_p1=cnt_p1+1
            if data.at[i,'Octant']==-1:
                cnt_n1=cnt_n1+1
            if data.at[i,'Octant']==2:
                cnt_p2=cnt_p2+1
            if data.at[i,'Octant']==-2:
                cnt_n2=cnt_n2+1
            if data.at[i,'Octant']==3:
                cnt_p3=cnt_p3+1
            if data.at[i,'Octant']==-3:
                cnt_n3=cnt_n3+1
            if data.at[i,'Octant']==4:
                cnt_p4=cnt_p4+1
            if data.at[i,'Octant']==-4:
                cnt_n4=cnt_n4+1
            i=i+1
        else:
            flag=True
    data.at[iter+1,'Octant ID']=str(prev)+"-"+str(mod*iter-1)
    if flag:
        data.at[iter+1,'Octant ID']=str(prev)+"-"+str(len(data))
    data.at[iter+1,'1']=cnt_p1
    data.at[iter+1,'-1']=cnt_n1
    data.at[iter+1,'2']=cnt_p2
    data.at[iter+1,'-2']=cnt_n2
    data.at[iter+1,'3']=cnt_p3
    data.at[iter+1,'-3']=cnt_n3
    data.at[iter+1,'4']=cnt_p4
    data.at[iter+1,'-4']=cnt_n4
    cnt_p1=0
    cnt_n1=0
    cnt_p2=0
    cnt_n2=0
    cnt_p3=0
    cnt_n3=0
    cnt_p4=0
    cnt_n4=0
    prev=mod*iter
    iter=iter+1

data.to_csv('octant_output.csv',index=False)

