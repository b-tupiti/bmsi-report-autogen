from helper import fetcher,sender
import generate 
import os

def runApp():

    # get attachments from daily email
    report_dump, sales = fetcher.fetch_mailattachments()

    # generate report excel file
    filename = generate.generate(sales, report_dump)

    # delete attachments
    dirname = 'input_files/'
    os.remove(dirname+report_dump)    
    os.remove(dirname+sales)

    # send report excel file to email list
    sender.send_report(filename)

if __name__ == "__main__":
    runApp()