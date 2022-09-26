# import openpyxl
# inp = openpyxl.load_workbook('input_octant_transition_identify.xlsx')
# sheet_input = inp.active

# output = openpyxl.Workbook()
# out = output.active
# mr = sheet_input.max_row
# mc = sheet_input.max_column
# for i in range (1, mr + 1):
# 	for j in range (1, mc + 1):
# 		# reading cell value from source excel file
# 		data = sheet_input.cell(row = i, column = j)

# 		# writing the read value to destination excel file
# 		out.cell(row = i, column = j).value = data.value

# output.save("output_octant_transition_identify.xlsx")

# output=openpyxl.load_workbook('output_octant_transition_identify.xlsx')
# sheet_output=output.active
# def octant_transition_count(mod=5000):
#     # sheet_output['F5']="hii"
#     sheet_output['E1']='U Avg'
#     sheet_output['E2']=f'=AVERAGE(B2:B{mr})'
#     sheet_output['F1']='V Avg'
#     sheet_output['F2']=f'=AVERAGE(C2:C{mr})'
#     sheet_output['G1']='W Avg'
#     sheet_output['G2']=f'=AVERAGE(D2:D{mr})'
#     sheet_output['H1']="U'=U-U avg"
#     sheet_output['I1']="V'=V-V avg"
#     sheet_output['J1']="W'=W-W avg"
#     for i in range (2, mr + 1):
#         for j in range (8, 11):
#             sheet_output.cell(row=i,column=j).value=sheet_output.cell(row=i,column=j-6).value-sheet_output.cell(row=2,column=j-3).value
#     output.save('output_octant_transition_identify.xlsx')
# from platform import python_version
# ver = python_version()

# if ver == "3.8.10":
#     print("Correct Version Installed")
# else:
#     print("Please install 3.8.10. Instruction are present in the GitHub Repo/Webmail. Url: https://pastebin.com/nvibxmjw")

# mod=5000
# octant_transition_count(mod)
# tut1 solution

import pandas as pd
import openpyxl
inp = openpyxl.load_workbook('input_octant_transition_identify.xlsx')
sheet_input = inp.active

output = openpyxl.Workbook()
out = output.active
mr = sheet_input.max_row
mc = sheet_input.max_column
for i in range (1, mr + 1):
	for j in range (1, mc + 1):
		# reading cell value from source excel file
		data = sheet_input.cell(row = i, column = j)

		# writing the read value to destination excel file
		out.cell(row = i, column = j).value = data.value

output.save("output_octant_transition_identify.xlsx")
data=pd.read_excel('output_octant_transition_identify.xlsx')
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

data.to_excel('output_octant_transition_identify.xlsx',index=False)

