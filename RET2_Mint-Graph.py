import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import plotly.graph_objects as go
import dash

a=int(input("Enter the maximum desired minutely throughput value: "))

file=("C:\\Users\\jacob\\Documents\\RET Files\\RET_2_0\\BT3.xlsx")
data=pd.read_excel(file)

#converting column data to date time format
data["Begin Time"]=pd.to_datetime(data["Begin Time"])
#extracting hour minute data from column
begin_time= data["Begin Time"].dt.strftime("%H:%M")

#hour minute list
mnt=[ ]
mng=[ ]

#minutely count list
mntc=[ ]
mngc=[ ]

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
        if ctr >= a:
               print("Minutely Throughput at ",start_time,": ", ctr)
               mng.append(start_time)
               mngc.append(ctr)
x=np.array(mng)
y=np.array(mngc)

def addlabels(x,y):
       for i in range(len(x)):
              plt.text(i,y[i],y[i], ha='center', va='bottom')

fig = go.Figure(data=go.Bar(x=mng, y=mngc, text=y, textposition='outside', width=0.4))
fig.update_layout(title_text=f"Minutely Throughput - Above {a} Shipments",title_x=0.5,title_font_size=15)
fig.write_html('minutely.html', auto_open=True)

report_title={"Minutely Throughput":mng, "Count":mngc}
report_data=pd.DataFrame(report_title)
print(report_data)

report_data.to_csv("C:\\Users\\jacob\\Documents\\RET Files\\RET_2_0\\Minutely_Data.csv",sep="\t")

#writing to file
with open('C:\\Users\\Jacob Thomas\\Documents\\Python Program\\Plot\\minthrptgh.txt', 'w') as f:
       for dm, dc in zip(mng,mngc):
             f.writelines('Minutely Throughput at ')
             f.writelines(str(dm))
             f.writelines(': ')
             f.writelines(str(dc))
             f.writelines('\n')
             #f.writeslines("{}\t{}\n".format(mnt,mntc))
f.close

#max_value = None

#for num in mntc:
#    if (max_value is None or num > max_value):
#        max_value = num

#print("The highest minutely throughput",max_value)        



x=np.array(mng)
y=np.array(mngc)
#z=[0,50,100,150,200,250,300]
#print(mng)
#print(mngc)
addlabels(x,y)
plt.title(f"Minutely Throughput - Above {a} Shipments", fontsize=15)
plt.yticks([])
plt.bar(x,y, width=0.4)
plt.show() 

