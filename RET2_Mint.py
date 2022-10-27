import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import plotly.graph_objects as go
import dash


file=("C:\\Users\\Jacob Thomas\\Documents\\Python Program\\My Programs\\BT.xlsx")
data=pd.read_excel(file)

#converting column data to date time format
data["Begin Time"]=pd.to_datetime(data["Begin Time"])
#extracting hour minute data from column
begin_time= data["Begin Time"].dt.strftime("%H:%M")

#hour minute list
mnt=[ ]

#miniutely count list
mntc=[ ]

start_time=""

#extracting data
for i in range (0, 24):
    for j in range (0,60):
        if (i < 10):
              if(j < 10): 
                    start_time = "0" + str(i) + ":0"+str(j)

              else:
                   start_time = "0" + str(i) + ":"+str(j)
        else:
              if(j < 10): 
                    start_time = str(i) + ":0"+str(j)
              else:
                    start_time = str(i) + ":"+str(j)
               
        mnt.append(start_time)
        ctr=0
        for x in begin_time:
             if start_time in str(x):
                   ctr += 1
        #print(ctr)                                      
        mntc.append(ctr)
            
max_value = None

for num in mntc:
    if (max_value is None or num > max_value):
        max_value = num

c1=0
c2=0
c3=0
for mc in mntc:
       if (mc>=230 and mc<240):
             c1+=1
       elif (mc>=240 and mc<250):
              c2+=1
       elif(mc>=250):
             c3+=1

print("Count of time more than 230",c1,'\n')
print("Count of time more than 240",c2,'\n')
print("Count of time more than 250",c3,'\n')

mnt.append('Maximum value:')
mntc.append(max_value)
#print('Maximum value:', max_value)

#writing to file
with open('C:\\Users\\Jacob Thomas\\Documents\\Python Program\\Plot\\minthrpt.txt', 'w') as f:
       for dm, dc in zip(mnt,mntc):
             f.writelines(str(dm))
             f.writelines('- ')
             f.writelines(str(dc))
             f.writelines('\n')
             #f.writeslines("{}\t{}\n".format(mnt,mntc))
f.close

#fig = go.Figure(data=go.Bar(x=mnt, y=mntc, text=y, textposition='outside', width=0.4))
#fig.update_layout(title_text="Per Minute Shipment Analysis",title_x=0.5,title_font_size=15)
#fig.write_html('minutely.html', auto_open=True)
