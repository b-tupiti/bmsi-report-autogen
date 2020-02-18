from datetime import datetime
import pandas as pd

def sales_to_df(lst):

    row = []
    table = []
    
    for obj in lst:
        arr = obj.split('|')
        arr[0] = datetime.strptime(arr[0],'%Y-%m-%d').date()
        arr[1] = int(arr[1])
        row.append(arr[0])
        row.append(arr[1])
        table.append(row)
        row=[]
    
    sales_df = pd.DataFrame(table,columns=['Date','Amount'])

    return sales_df

def FCA_list_to_df(FCA_dl):
    
    row = []
    table = []
    for strobj in FCA_dl:
        obj_arr = strobj.split('|')
        for i in range(len(obj_arr)):
            if not obj_arr[i]:
                obj_arr[i] = '0'
        
        obj_arr[0] = datetime.strptime(obj_arr[0], '%Y-%m-%d').date()
        obj_arr[1] = int(obj_arr[1])
        
        for obj in obj_arr:
            row.append(obj)
        table.append(row)
        row = []

    fca_df = pd.DataFrame(table,columns=['Date','Number of FCA'])
    return fca_df

def ASS_list_to_df(ASS_dl):

    row = []
    table = []
    
    for strobj in ASS_dl:
        obj_arr = strobj.split('|')
        for i in range(len(obj_arr)):
            if not obj_arr[i]:
                obj_arr[i] = '0'
        
        obj_arr[0] = datetime.strptime(obj_arr[0], '%Y-%m-%d').date()
        obj_arr[1] = int(obj_arr[1])
        
        for obj in obj_arr:
            row.append(obj)
        table.append(row)
        row = []

    ass_df = pd.DataFrame(table,columns=['Date','Customer Count'])
    
    return ass_df



def CDB_list_to_df(CDB_dl):
    row = []
    table = []
    for strobj in CDB_dl:
        obj_arr = strobj.split('|')
        for i in range(len(obj_arr)):
            if not obj_arr[i]:
                obj_arr[i] = '0'       
        
        obj_arr[0] = datetime.strptime(obj_arr[0], '%Y-%m-%d').date()
        for i in range(1,16):
            obj_arr[i] = int(obj_arr[i])
            
        for obj in obj_arr:
            row.append(obj)
        table.append(row)
        row = []
    
    cdb_df = pd.DataFrame(table,columns=['Date','ON_NET_CALL','ON_NET_MIN','ON_NET_COST',
                                'OFF_NET_CALL','OFF_NET_MIN','OFF_NET_COST',
                                'TEL_CALL','TEL_MIN','TEL_COST',
                                'INT_CALL','INT_MIN','INT_COST',
                                'OTH_CALL','OTH_MIN','OTH_COST'])
    
    return cdb_df

def CDNB_list_to_df(CDNB_dl):
    row = []
    table = []
    for strobj in CDNB_dl:
        obj_arr = strobj.split('|')
        for i in range(len(obj_arr)):
            if not obj_arr[i]:
                obj_arr[i] = '0'
        
        obj_arr[0] = datetime.strptime(obj_arr[0],'%Y-%m-%d').date()
        for i in range(1,11):
            obj_arr[i] = int(obj_arr[i])
            
        for obj in obj_arr:
            row.append(obj)
        table.append(row)
        row=[]
       
    cdnb_df = pd.DataFrame(table,columns=['Date','Nbr of Calls','Minutes',
                                'Nbr of Calls','Minutes',
                                'Nbr of Calls','Minutes',
                                'Nbr of Calls','Minutes',
                                'Nbr of Calls','Minutes'])
    
    return cdnb_df

def SMSDB_list_to_df(SMSDB_dl):
    row = []
    table = []
    for strobj in SMSDB_dl:
        obj_arr = strobj.split('|')
        for i in range(len(obj_arr)):
            if not obj_arr[i]:
                obj_arr[i] = '0'
        
        obj_arr[0] = datetime.strptime(obj_arr[0],'%Y-%m-%d').date()
        for i in range(1,9):
            obj_arr[i] = int(obj_arr[i])

        for obj in obj_arr:
            row.append(obj)
        table.append(row)
        row=[]
   
    smsdb_df = pd.DataFrame(table,columns=["date","Nbr of SMS","Face Value($)",
                                 "Nbr of SMS","Face Value($)",
                                 "Nbr of SMS","Face Value($)",
                                 "Nbr of SMS","Face Value($)"])
    
    return smsdb_df

def SMSDNB_list_to_df(SMSDNB_dl):
    row = []
    table = []
    for strobj in SMSDNB_dl:
        obj_arr = strobj.split('|')
        for i in range(len(obj_arr)):
            if not obj_arr[i]:
                obj_arr[i] = '0'
        
        obj_arr[0] = datetime.strptime(obj_arr[0],'%Y-%m-%d').date()
        for i in range(1,5):
            obj_arr[i] = int(obj_arr[i])

        for obj in obj_arr:
            row.append(obj)
        table.append(row)
        row=[]
   
    smsdnb_df = pd.DataFrame(table,columns=["date","No of SMS Onnet",
                                "No of SMS Offnet","No of SMS Intnl",
                                "No of SMS Other"])
    
    return smsdnb_df

def GPRSBNB_list_to_df(GPRSBNB_dl):
    row = []
    table = []
    for strobj in GPRSBNB_dl:
        obj_arr = strobj.split('|')
        for i in range(len(obj_arr)):
            if not obj_arr[i]:
                obj_arr[i] = '0'
        
        obj_arr[0] = datetime.strptime(obj_arr[0],'%Y-%m-%d').date()
        for i in range(1,9):
            obj_arr[i] = int(obj_arr[i])

        for obj in obj_arr:
            row.append(obj)
        table.append(row)
        row=[]
   
    gprsbnb_df = pd.DataFrame(table,columns=["date","Nbr of Events",
                                "Total MBytes","Face Value (SBD)",
                                "Nbr of Subs","Nbr of Events",
                                "Total Mbytes","Total SBD",
                                "NBR of SUBS"])
    
    return gprsbnb_df

def SD_list_to_df(SD_dl):
    row = []
    table = []
    for strobj in SD_dl:
        obj_arr = strobj.split('|')
        for i in range(len(obj_arr)):
            if not obj_arr[i]:
                obj_arr[i] = '0'
        
        obj_arr[0] = datetime.strptime(obj_arr[0],'%Y-%m-%d').date()
        
        float_cols = [2,5,8,11,14,17,20,23,26,29,32,35,38,41,44,47,50] 
        for i in range(1,52):
            if i in float_cols:
                obj_arr[i] = float(obj_arr[i])
            else:
                obj_arr[i] = int(obj_arr[i])

        for obj in obj_arr:
            row.append(obj)
        table.append(row)
        row=[]
   
    sd_df = pd.DataFrame(table,columns=['Date','Count','Amount','Unique Users',
                                'Count','Amount','Unique Users',
                                'Count','Amount','Unique Users',
                                'Count','Amount','Unique Users',
                                'Count','Amount','Unique Users',
                                'Count','Amount','Unique Users',
                                'Count','Amount','Unique Users',
                                'Count','Amount','Unique Users',
                                'Count','Amount','Unique Users',
                                'Count','Amount','Unique Users',
                                'Count','Amount','Unique Users',
                                'Count','Amount','Unique Users',
                                'Count','Amount','Unique Users',
                                'Count','Amount','Unique Users',
                                'Count','Amount','Unique Users',
                                'Count','Amount','Unique Users',
                                'Count','Amount','Unique Users'])
    
    return sd_df

def RbS_list_to_df(RbS_dl):
    row = []
    table = []
    for strobj in RbS_dl:
        obj_arr = strobj.split('|')
        for i in range(len(obj_arr)):
            if not obj_arr[i]:
                obj_arr[i] = '0'
        
        obj_arr[0] = datetime.strptime(obj_arr[0],'%Y-%m-%d').date()
        
        for i in range(1,4):
            if i==2:
                obj_arr[i] = float(obj_arr[i])
            else:
                obj_arr[i] = int(obj_arr[i])

        for obj in obj_arr:
            row.append(obj)
        table.append(row)
        row=[]
   
    rbs_df = pd.DataFrame(table,columns=['Date','NBR_RECHARGE','TOTAL_RECHARGE','NBR_SUBS'])
    
    return rbs_df

def BT_list_to_df(BT_dl):
    row = []
    table = []
    for strobj in BT_dl:
        obj_arr = strobj.split('|')
        for i in range(len(obj_arr)):
            if not obj_arr[i]:
                obj_arr[i] = '0'
        
        obj_arr[0] = datetime.strptime(obj_arr[0],'%Y-%m-%d').date()
        
        float_cols = [3,4]
        for i in range(1,5):
            if i in float_cols:
                obj_arr[i] = float(obj_arr[i])
            else:
                obj_arr[i] = int(obj_arr[i])

        for obj in obj_arr:
            row.append(obj)
        table.append(row)
        row=[]
   
    bt_df = pd.DataFrame(table,columns=['Date','NBR of Transaction','Nbr of Subs','SBD TRANSFERRED','TRANSACTION FEE'])
    
    return bt_df

def RM_list_to_df(RM_dl):
    row = []
    table = []
    for strobj in RM_dl:
        obj_arr = strobj.split('|')
        for i in range(len(obj_arr)):
            if not obj_arr[i]:
                obj_arr[i] = '0'
        
        obj_arr[0] = datetime.strptime(obj_arr[0],'%Y-%m-%d').date()
        
        float_cols = [2,5,8,11,14,17] 
        for i in range(1,19):
            if i in float_cols:
                obj_arr[i] = float(obj_arr[i])
            else:
                obj_arr[i] = int(obj_arr[i])

        for obj in obj_arr:
            row.append(obj)
        table.append(row)
        row=[]
   
    rm_df = pd.DataFrame(table,columns=['Date','Nbr of Recharge','Recharge Value(SBD)','Subs',
                                'Nbr of Recharge','Recharge Value(SBD)','Subs',
                                'Nbr of Recharge','Recharge Value(SBD)','Subs',
                                'STAF_NBR','Rech(SBD)','Staff Subs',
                                'CUG_NBR','CUG_SBD','CUG SUBS',
                                'POB_NBR','POB_SBD','POB SUBS'])
    
    return rm_df


def IM_list_to_df(IM_dl):
    row = []
    table = []
    for strobj in IM_dl:
        obj_arr = strobj.split('|')
        for i in range(len(obj_arr)):
            if not obj_arr[i]:
                obj_arr[i] = '0'
        
        obj_arr[0] = datetime.strptime(obj_arr[0],'%Y-%m-%d').date()
        
        for i in range(1,7):
            obj_arr[i] = int(obj_arr[i])

        for obj in obj_arr:
            row.append(obj)
        table.append(row)
        row=[]
   
    im_df = pd.DataFrame(table,columns=['Date','Minutes','Subs',
                                 'Minutes','Subs',
                                 'Minutes','Subs'])
    return im_df

def ERR_list_to_df(ERR_dl):
    row = []
    table = []
    for strobj in ERR_dl:
        obj_arr = strobj.split('|')
        for i in range(len(obj_arr)):
            if not obj_arr[i]:
                obj_arr[i] = '0'
        
        obj_arr[0] = datetime.strptime(obj_arr[0],'%Y-%m-%d').date()
        obj_arr[1] = float(obj_arr[1])

        for obj in obj_arr:
            row.append(obj)
        table.append(row)
        row=[]
   
    err_df = pd.DataFrame(table,columns=['Date','Amount'])
    return err_df

def CP_list_to_df(CP_dl):
    row = []
    table = []
    for strobj in CP_dl:
        obj_arr = strobj.split('|')
        for i in range(len(obj_arr)):
            if not obj_arr[i]:
                obj_arr[i] = '0'
        
        obj_arr[0] = datetime.strptime(obj_arr[0],'%Y-%m-%d').date()
        for i in range(2,5):
            if i == 3:
                obj_arr[i] = float(obj_arr[i])
            else:
                obj_arr[i] = int(obj_arr[i])
        

        for obj in obj_arr:
            row.append(obj)
        table.append(row)
        row=[]
   
    cp_df = pd.DataFrame(table,columns=['Date','Product_Desc','Count',
                                 'Revenue','Unique Users'])
    return cp_df

def oCUG_GSD_list_to_df(oCUG_GSD_dl):
    row = []
    table = []
    for strobj in oCUG_GSD_dl:
        obj_arr = strobj.split('|')
        for i in range(len(obj_arr)):
            if not obj_arr[i]:
                obj_arr[i] = '0'
        
        obj_arr[0] = datetime.strptime(obj_arr[0],'%Y-%m-%d').date()
        for i in range(2,5):
            if i == 2:
                obj_arr[i] = float(obj_arr[i])
            else:
                obj_arr[i] = int(obj_arr[i])
        

        for obj in obj_arr:
            row.append(obj)
        table.append(row)
        row=[]
   
    ocug_gsd_df = pd.DataFrame(table,columns=['Date','Destination','Revenue',
                                 'unique users','calls'])
    return ocug_gsd_df


def RVbD_list_to_df(RVbD_dl):
    row = []
    table = []
    for strobj in RVbD_dl:
        obj_arr = strobj.split('|')
        for i in range(len(obj_arr)):
            if not obj_arr[i]:
                obj_arr[i] = '0'
        
        obj_arr[0] = datetime.strptime(obj_arr[0],'%Y-%m-%d').date()
        for i in range(1,5):
            if i == 3:
                obj_arr[i] = float(obj_arr[i])
            else:
                obj_arr[i] = int(obj_arr[i])
        

        for obj in obj_arr:
            row.append(obj)
        table.append(row)
        row=[]
   
    rvbd_df = pd.DataFrame(table,columns=['Entry Date','Denomination','Unique Users', 'Recharge Amount','Recharge Count'])
    return rvbd_df