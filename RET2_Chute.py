import pandas as pd
import numpy as np
from matplotlib import pyplot as plt
import seaborn as sns
import plotly.figure_factory as ff
import dash

file=("C:\\Users\\jacob\\Downloads\\123.xls")
chulist=[('ch0',[])]
data=pd.read_excel(file)
# print(data.shape[0])
table_size=data.shape[0]
chulist=[]
for i in range (0,table_size):
        temp=[(data["Chute Name"],data["X coordinates"],data["X coordinates"],)]

chutel[0]

chute_data=data['Chute Name']

def showdata(z_data,filename):
    col = list(z_data.columns)
    ind = list(z_data.index)
    fig = go.Figure(
        data=[go.Surface(z=z_data.values, x=col, y=ind, colorscale='Viridis',)])
    noaxis = dict(showbackground=False,
                  showgrid=False,
                  showline=False,
                  showticklabels=False,
                  ticks='',
                  title='',
                  zeroline=False)
    fig.update_layout( autosize=False,
                      width=750, height=750, scene_aspectratio=dict(x=1, y=1.2, z=0.5),
                      scene_aspectmode='manual',
                      margin=dict(l=65, r=50, b=65, t=90), scene_camera_eye=dict(x=1, y=0, z=2),yaxis_title=None)
    fig.update_traces(showscale=False)
    # fig.show()
    fig.write_image(filename,scale=5)
    # fig.write_html(loc,include_plotlyjs="cdn", full_html=False)
    data = fig.to_html(include_plotlyjs="cdn", full_html=False)
    gdata=sepratehtml(data)
    return gdata

# print(chute_data)

# chute_name=""
# chnm=[ ]
# chcn=[ ]
# chcdx=[ ]
# chcdy=[ ]

# for i in range(1,519):
#          if (i<10):
#                  chute_name="CH"+"0"+"0"+str(i)
#                  #print(chute_name)
              
#          else:
#                  if(i<100):
#                          chute_name="CH"+"0"+str(i)
#                           #print(chute_name)
                    
#                  else:
#                           chute_name="CH"+str(i)
#                           #print(chute_name)
                          
#          chnm.append(chute_name)       
#          cn=0
#          for j in chute_data:
#                  if chute_name in str(j):
#                          cn +=1
#          chcn.append(cn)
#                  #print(cn)

# print(chnm)
# print(chcn)

# for xc in range(96, 151):
#          chcdy.append(xc)
# print(chcdx)

# for yc in range(86, 186):
#          chcdy.append(yc)
# print(chcdy)

# z=[chcn]
# z_text= [chnm]

# fig = ff.create_annotated_heatmap(chcdx, x=chcdx, y=chcdx, annotation_text=z_text, colorscale='Virdis')
# fig.write_html('chute_conc.html', auto_open=True)

# with open('C:\\Users\\Jacob Thomas\\Documents\\Python Program\\Plot\\Chute_conc.txt', 'w') as f:
#        for dm, dc in zip(chnm,chcn):
#              f.writelines(str(dm))
#              f.writelines('- ')
#              f.writelines(str(dc))
#              f.writelines('\n')
#              #f.writeslines("{}\t{}\n".format(mnt,mntc))
# f.close
