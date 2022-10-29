import csv
import pandas as pd
from datetime import datetime
start_time = datetime.now()

def attendance_report():
    try:
        inp_file = pd.read_csv('input_attendance.csv')
        inp = inp_file.fillna("20010000 Random")
    except:
        print("File not found")
    
    
    try:
        rollno_inp=pd.read_csv('input_registered_students.csv')
    except:
        print('File containing name of all students is missing!')
    

    # mr = sheet_input.max_row

    mc=sum(1 for row in open("input_registered_students.csv"))
    mc_consolidated=sum(1 for row in open("input_attendance.csv"))
    total_dates=list()
    for i in range(0,mc_consolidated-1):
        if inp.at[i,'Timestamp'].split()[0] not in total_dates:
            total_dates.append(inp.at[i,'Timestamp'].split()[0])
    # max_att=0
    # max_roll=""
    fileName_consolidated=".\myoutput\\attendance_report_consolidated.csv"
    fileName_duplicate=".\myoutput\\attendance_report_duplicate.csv"
    with open(fileName_consolidated,'w',newline='') as output_file:
        writer=csv.writer(output_file)
        header = ['Roll', 'Name', 'total_lecture_taken', 'attendance_count_actual','attendance_count_fake','attendance_count_absent','Percentage']
        writer.writerow(header)
    with open(fileName_duplicate,'w',newline='') as output_file:
        writer=csv.writer(output_file)
        header = ["Timestamp","Roll","Name","Total count of attendance on that day"]
        writer.writerow(header)
    attend_consolidated=pd.read_csv(fileName_consolidated)
    attend_duplicate=pd.read_csv(fileName_duplicate)
    duplicate_index=0
    for i in range(0,mc-1):
        rollno=rollno_inp.at[i,'Roll No']
        t_lec,t_lec_act,t_lec_fake,t_lec_abs,percent=len(total_dates),0,0,0,0
        for j in range(0,mc_consolidated-1):
            unique_dates=list()
            if inp.at[j,'Attendance'].split()[0]==rollno:
                day=inp.at[j,'Timestamp'].split()[0][0:2]
                month=inp.at[j,'Timestamp'].split()[0][3:5]
                year=inp.at[j,'Timestamp'].split()[0][6:10]
                date=datetime.strptime(f'{year}-{month}-{day}', "%Y-%m-%d").date()
                day_name=date.strftime("%A")

                if day_name=='Monday' or day_name=='Thursday':
                    if (inp.at[j,'Timestamp'].split()[1].split(':')[0] == '14'):
                        if inp.at[j,'Timestamp'].split()[0] not in unique_dates:
                            unique_dates.append(inp.at[j,'Timestamp'].split()[0])
                            t_lec_act+=1
                        else:
                            attend_duplicate.at[duplicate_index,'Timestamp']=inp.at[j,'Timestamp']
                            attend_duplicate.at[duplicate_index,'Roll']=rollno
                            attend_duplicate.at[duplicate_index,'Name']=rollno_inp.at[i,'Name']
                            attend_duplicate.at[duplicate_index,'Total count of attendance on that day']=
                    else:
                        t_lec_fake+=1
                else: t_lec_fake+=1
        fileName=".\myoutput\\"+rollno+'.csv'
        with open(fileName,'w',newline='') as output_file:
            writer=csv.writer(output_file)
            header = ['Roll', 'Name', 'total_lecture_taken', 'attendance_count_actual','attendance_count_fake','attendance_count_absent','Percentage']
            writer.writerow(header)
        out=pd.read_csv(fileName)
        out.at[0,'Roll']=rollno
        attend_consolidated.at[i,'Roll']=rollno
        out.at[0,'Name']=rollno_inp.at[i,'Name']
        attend_consolidated.at[i,'Name']=rollno_inp.at[i,'Name']
        out.at[0,'total_lecture_taken']=t_lec
        attend_consolidated.at[i,'total_lecture_taken']=t_lec
        out.at[0,'attendance_count_actual']=t_lec_act
        attend_consolidated.at[i,'attendance_count_actual']=t_lec_act
        # if t_lec_act>max_att:
        #     max_att=t_lec_act
        #     max_roll=rollno
        out.at[0,'attendance_count_fake']=t_lec_fake
        attend_consolidated.at[i,'attendance_count_fake']=t_lec_fake
        out.at[0,'attendance_count_absent']= t_lec_abs = t_lec-t_lec_act
        attend_consolidated.at[i,'attendance_count_absent']= t_lec_abs = t_lec-t_lec_act
        out.at[0,'Percentage']=round((t_lec_act/t_lec)*100,2)
        attend_consolidated.at[i,'Percentage']=round((t_lec_act/t_lec)*100,2)
        out.to_csv(fileName,index=False)
    attend_consolidated.to_csv(fileName_consolidated,index=False)
    # print(max_roll,max_att)
# ,,,,,, (attendance_count_actual/total_lecture_taken) 2 digit decimal 

from platform import python_version
ver = python_version()

if ver == "3.8.10":
    print("Correct Version Installed")
else:
    print("Please install 3.8.10. Instruction are present in the GitHub Repo/Webmail. Url: https://pastebin.com/nvibxmjw")


attendance_report()




#This shall be the last lines of the code.
end_time = datetime.now()
print('Duration of Program Execution: {}'.format(end_time - start_time))
