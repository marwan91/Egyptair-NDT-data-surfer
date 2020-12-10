
# coding: utf-8

# In[3]:


import wx
import wx.grid
import pandas as pd
import numpy as np
import glob


######################## Data Loading ############################

method_files=[]

try:
    xray_files = glob.glob('./xray/*.xlsx')
    method_files+=['X-ray']*len(xray_files)
except:
    xray_files=[]
    print('no xray files')
try:
    thermography_files = glob.glob('./thermography/*.xlsx')
    method_files+=['Thermography']*len(thermography_files)
except:
    thermography_files = []
    print('no thermography files')
try:
    ultrasonic_files = glob.glob('./ultrasonic/*.xlsx')
    method_files+=['Ultrasonic']*len(ultrasonic_files)
except:
    ultrasonic_files = []
    print('no ultrasonic files')
try:
    eddycurrent_files = glob.glob('./eddycurrent/*.xlsx')
    method_files+=['Eddy Current']*len(eddycurrent_files)
except:
    ultrasonic_files = []
    print('no eddycurrent files')
    

all_files_list = xray_files+thermography_files+ultrasonic_files+eddycurrent_files


desc_list=[]
pn_list=[]
sn_list = []
loc_list = []
qty_list = []
method_list = []
item_list = []

desc_string = "Unnamed: 1"
pn_string = "Unnamed: 3"
sn_string = "Unnamed: 4"
loc_string = "Unnamed: 6"
qty_string = "Unnamed: 7"
item_string = "Unnamed: 0"


for j in range(len(all_files_list)):
    
   
    data = pd.read_excel (all_files_list[j],None)
    
    for key in data:
        df = pd.DataFrame(data[key], columns= ['Unnamed: 0','Unnamed: 1','Unnamed: 2','Unnamed: 3','Unnamed: 4','Unnamed: 5','Unnamed: 6','Unnamed: 7','Unnamed: 8'])
    
        if ('DESCRIPITON' in df["Unnamed: 2"].tolist()) and ('P/N' in df["Unnamed: 4"].tolist()) and ('S/N' in df["Unnamed: 5"].tolist()):
            desc_string = "Unnamed: 2"
            pn_string = "Unnamed: 4"
            sn_string = "Unnamed: 5"
            loc_string = "Unnamed: 7"
            qty_string = "Unnamed: 8"
            item_string = "Unnamed: 1"
        
        elif ('DESCRIPITON' in df["Unnamed: 1"].tolist()) and ('P/N' in df["Unnamed: 3"].tolist()) and ('S/N' in df["Unnamed: 4"].tolist()):
            desc_string = "Unnamed: 1"
            pn_string = "Unnamed: 3"
            sn_string = "Unnamed: 4"
            loc_string = "Unnamed: 6"
            qty_string = "Unnamed: 7"
            item_string = "Unnamed: 0"
            
        elif ('DESCRIPITON' in df["Unnamed: 0"].tolist()) and ('P/N' in df["Unnamed: 2"].tolist()) and ('S/N' in df["Unnamed: 3"].tolist()): 
        
            desc_string = "Unnamed: 0"
            pn_string = "Unnamed: 2"
            sn_string = "Unnamed: 3"
            loc_string = "Unnamed: 5"
            qty_string = "Unnamed: 6"
            item_string = "Unnamed: 0"
            
            
            
        new_desc_list = df[desc_string].tolist()
        desc_list+=new_desc_list
        new_pn_list = df[pn_string].tolist()
        pn_list+=new_pn_list
        new_sn_list = df[sn_string].tolist()
        sn_list+= new_sn_list
    
        new_loc_list = df[loc_string].tolist()
        loc_list+= new_loc_list
    
        new_qty_list = df[qty_string].tolist()
        qty_list+= new_qty_list
        
        new_item_list = df[item_string].tolist()
        item_list+= new_item_list
        
        method_list+=[method_files[j]]*len(new_desc_list)



############################# Window Initialization ################


app = wx.App()
frame = wx.Frame(parent=None, title='Egyptair NDT Department')
frame.SetIcon(wx.Icon("logo.png"))
panel = wx.Panel(frame,pos=(0,0), size = (2000,200))
text_ctrl = wx.TextCtrl(panel, pos=(5, 25))
my_btn = wx.Button(panel, label='Search', pos=(5, 60))
data_panel = wx.Panel(frame, pos = (0,220), size =(1000,500))
guide_text = wx.StaticText(panel, pos = (5,5), label = 'Search by Part Number or Serial Number')


######################## Data Display ################################


new_desc_list=[]
new_pn_list=[]
new_sn_list=[]
new_loc_list = []
new_qty_list = []
new_method_list = []
new_item_list = []

for i in range(len(desc_list)): # Clear unwanted elements
    
    desc = desc_list[i]
    pn = pn_list[i]
    sn = sn_list[i]
    loc = loc_list[i]
    qty = qty_list[i]
    method = method_list[i]
    item = item_list[i]
    
    if (desc=='DESCRIPITON' and pn == 'P/N' and sn=='S/N'):
        continue
    if (pd.isna(desc) and pd.isna(pn) and pd.isna(sn)):
        continue
    if (pd.isna(desc) and not pd.isna(item)):
        desc=item
    
    new_desc_list.append(desc)
    new_pn_list.append(pn)
    new_sn_list.append(sn)
    new_loc_list.append(loc)
    new_qty_list.append(qty)
    new_method_list.append(method)
    

desc_list = new_desc_list
pn_list = new_pn_list
sn_list = new_sn_list
loc_list = new_loc_list
qty_list = new_qty_list
method_list = new_method_list

mygrid = wx.grid.Grid(data_panel, pos = (0,0), size = (1000,400))
mygrid.CreateGrid(1 ,6)

mygrid.SetColLabelValue(0, "DESCRIPTION")
mygrid.SetColLabelValue(1, "Part Number")
mygrid.SetColLabelValue(2, "Serial Number")
mygrid.SetColLabelValue(3, "Location")
mygrid.SetColLabelValue(4, "Quantity")
mygrid.SetColLabelValue(5, "Method")


def search_clicked(event):
    query = str(text_ctrl.GetValue())
    
    desc_result = []
    pn_result = []
    sn_result = []
    loc_result = []
    qty_result = []
    method_result = []
    
    for i in range(len(desc_list)): # Clear unwanted elements
        
        if query == "":
            return
    
        if (query in   str(pn_list[i])) or (query in str(sn_list[i])) :
        
            desc_result.append(desc_list[i])
            pn_result.append(pn_list[i])
            sn_result.append(sn_list[i])
            loc_result.append(loc_list[i])
            qty_result.append(qty_list[i])
            method_result.append(method_list[i])
    
    mygrid.ClearGrid()
    try:
        mygrid.DeleteRows(0, mygrid.GetNumberRows())
    except:
        pass
    mygrid.SetColLabelValue(0, "DESCRIPTION")
    mygrid.SetColLabelValue(1, "Part Number")
    mygrid.SetColLabelValue(2, "Serial Number")
    mygrid.SetColLabelValue(3, "Location")
    mygrid.SetColLabelValue(4, "Quantity")
    mygrid.SetColLabelValue(5, "Method")
    mygrid.AppendRows( len(desc_result))
    
    
    for i in range(len(desc_result)): #Rows
        desc = desc_result[i]
        pn = pn_result[i]
        sn = sn_result[i]
        loc = loc_result[i]
        qty = qty_result[i]
        method = method_result[i]
    
        mygrid.SetCellValue (i,0,str(desc))
        mygrid.SetCellValue (i,1,str(pn) )
        mygrid.SetCellValue (i,2,str(sn) )
        mygrid.SetCellValue (i,3,str(loc) )
        mygrid.SetCellValue (i,4,str(qty) )
        mygrid.SetCellValue (i,5,str(method) )
    mygrid.AutoSizeColumns(True)
    
    
############################### Binding ######################

my_btn.Bind(wx.EVT_BUTTON,search_clicked) 
    
####################################################################

frame.Show()
app.MainLoop()
del app

