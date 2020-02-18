from datetime import datetime, timedelta

def generate_filename():

    default_str = 'SI_Daily_Traffic_Recharge_Report['
    filename = ''

    today = datetime.now().strftime('%d')
    first_of_the_month = '1'
    if today == first_of_the_month:

        yesterday = datetime.strftime(datetime.now() - timedelta(1),'%d')
        yester_month = datetime.strftime(datetime.now() - timedelta(1),'%b')
        yester_year = datetime.strftime(datetime.now() - timedelta(1),'%Y')
        filename = default_str + yester_month + '_' + first_of_the_month + '-' + yesterday + '_' + yester_year + ']' 


    else:
        month = datetime.now().strftime('%b')
        year = datetime.now().strftime('%Y')
        period = month + '_' + first_of_the_month + '-' + today + '_' + year

        filename =  default_str + period + ']'
    return filename
