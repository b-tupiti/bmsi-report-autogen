from datetime import datetime

cur_month = datetime.now().month
today = datetime.now()
today = str(today).split(' ')[0]
today_num = today.split('-')[2]

def filter_for_month(lst):
    if today_num == '01':
        newlst = []
        for line in lst:
            # get everything from the previous month
            # get previous month
            if cur_month == 1: # if month is january, then set prev to december
                report_month = 12
                # get all the data from december
                datestr = line.split('|')[0]
                date = datetime.strptime(datestr, "%Y-%m-%d")
                month = date.month
                if month == report_month:
                    newlst.append(line)
            else:
                report_month = cur_month - 1
                datestr = line.split('|')[0]
                date = datetime.strptime(datestr, "%Y-%m-%d")
                month = date.month
                if month == report_month:
                    newlst.append(line)
    else:    
        newlst = []
        for line in lst:
            datestr = line.split('|')[0]
            date = datetime.strptime(datestr, "%Y-%m-%d")
            month = date.month
            if month == cur_month:
                if today == datestr:
                    break
                else:
                    newlst.append(line)
                
    
    return newlst

