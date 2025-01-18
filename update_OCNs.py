import configparser
import json
import re
import requests
import sys
import time
import xlrd
import xlwt
import xlutils.copy
from threading import Thread
from queue import Queue
from time import sleep
from ratelimit import limits, sleep_and_retry
from lxml import etree

# Read config file
config = configparser.ConfigParser()
config.read('local_settings.ini')
key = config['Alma Bibs R/W']['key']

num_worker_threads = 15
work_queue = Queue()
output_queue = Queue()

bibID_col_index = 1
newOCN_col_index = 2
getbib_col_index = 3
existingOCN_col_index = 4
existingVihartem_col_index = 5
existing035s_col_index = 6
putbib_col_index = 7
updatedOCN_col_index = 8
updatedVihartem_col_index = 9
updated035s_col_index = 10

@sleep_and_retry
@limits(calls=15, period=1)
def api_request(type, bib, record=None):
    if type == 'get':
        headers = {'accept':'application/xml'}
        response = requests.get('https://api-na.hosted.exlibrisgroup.com/almaws/v1/bibs/'+bib['bibID']+'?apikey='+key, headers=headers)
        return response
    if type == 'put':
        headers = {'accept':'application/xml', 'Content-Type':'application/xml'}
        response = requests.put('https://api-na.hosted.exlibrisgroup.com/almaws/v1/bibs/'+bib['bibID']+'?validate=false&override_warning=true&override_lock=true&stale_version_check=false&check_match=false&apikey='+key, headers=headers, data=record)
        return response

def prettyprint(element, **kwargs):
    xml = etree.tostring(element, pretty_print=True, **kwargs)
    print(xml.decode(), end='')

def worker():
    while True:
        bib = work_queue.get()
        output = []
        get_response = api_request('get', bib)
        output.append((bib['row'], getbib_col_index, get_response.status_code))
        
        if get_response.status_code == 200:
            bib_tree = etree.fromstring(get_response.content)
            
            # get content of all 035s in a list
            record_datafield_035s = bib_tree.xpath('//record/datafield[@tag="035"]/subfield')
            old_list_035a = []
            for i in record_datafield_035s:
                old_list_035a.append(i.text)
            # print('All 035s:')
            # print(old_list_035a)
            output.append((bib['row'], existing035s_col_index, ';'.join(old_list_035a)))
            # print()
            
            ocn_035 = ''
            vihartem_035 = ''
            
            ocn_035 = bib_tree.xpath('//record/datafield[@tag="035"]/subfield[contains(text(), "(OCoLC)")]')
            if len(ocn_035) == 1:
                # print('Old 035 OCoLC: ' + ocn_035[0].text)
                output.append((bib['row'], existingOCN_col_index, ocn_035[0].text))
                OCNnum = re.sub(r'^\(OCoLC\)(.*)$', r'\1', ocn_035[0].text)
                ocn_035[0].text = '(OCoLC)' + bib['newOCN']
                # print('New 035 OCoLC: ' + ocn_035[0].text)
            
            vihartem_035 = bib_tree.xpath('//record/datafield[@tag="035"]/subfield[contains(text(), "(ViHarT-EM)")]')
            if len(vihartem_035) == 1:
                # print('Old 035 ViHarT-EM: ' + vihartem_035[0].text)
                output.append((bib['row'], existingVihartem_col_index, vihartem_035[0].text))
                VihartemNum = re.sub(r'^\(ViHarT-EM\)(.*)$', r'\1', vihartem_035[0].text)
                
                # check for ViHarT-EM 035; check for whether it includes currentOCN; if so, update to newOCN
                # variable OCNnum is still the old OCLC number, so it's ok to use for comparison
                if OCNnum in VihartemNum:
                    vihartem_035[0].text = vihartem_035[0].text.replace(OCNnum, bib['newOCN'])
                # print('New 035 ViHarT-EM: ' + vihartem_035[0].text)
                # print()
            
            new_list_035a = []
            for i in record_datafield_035s:
                new_list_035a.append(i.text)
            # print('All 035s, updated:')
            # print(new_list_035a)
            
            put_response = api_request('put', bib, etree.tostring(bib_tree))
            output.append((bib['row'], putbib_col_index, put_response.status_code))
            # print(put_response.reason)
            if put_response.status_code == 400:
                output.append((bib['row'], updatedOCN_col_index, put_response.text))
            if put_response.status_code == 200:
                output.append((bib['row'], updatedOCN_col_index, ocn_035[0].text))
                if len(vihartem_035) == 1:
                    output.append((bib['row'], updatedVihartem_col_index, vihartem_035[0].text))
                output.append((bib['row'], updated035s_col_index, ';'.join(new_list_035a)))
            
        output_queue.put(output)
        print(bib['row'], bib['bibID'])
        work_queue.task_done()
        

def out_worker(book_in, input):
    # Copy spreadsheet for output
    book_out = xlutils.copy.copy(book_in)
    sheet_out = book_out.get_sheet(0)
    
    # Add new column headers
    sheet_out.write(0,getbib_col_index,'GET Bib')
    sheet_out.write(0,existingOCN_col_index,'Existing 035 OCoLC')
    sheet_out.write(0,existingVihartem_col_index,'Existing 035 ViharT-EM')
    sheet_out.write(0,existing035s_col_index,'Existing 035s')
    sheet_out.write(0,putbib_col_index,'PUT Bib')
    sheet_out.write(0,updatedOCN_col_index,'Updated 035 OCoLC')
    sheet_out.write(0,updatedVihartem_col_index,'Updated 035 ViHarT-EM')
    sheet_out.write(0,updated035s_col_index,'Updated 035s')
    
    while True:
        bibs = output_queue.get()
        for bib in bibs:
            sheet_out.write(bib[0], bib[1], bib[2])
        book_out.save(input+'_results.xls')
        output_queue.task_done()

def main(input):
    st = time.localtime()
    start_time = time.strftime("%H:%M:%S", st)
   
    # Read spreadsheet
    book_in = xlrd.open_workbook(input)
    sheet1 = book_in.sheet_by_index(0) #get first sheet
    
    Thread(target=out_worker, args=(book_in, input,), daemon=True).start()
    for i in range(num_worker_threads):
        Thread(target=worker, daemon=True).start()
    
    
    for row in range(1, sheet1.nrows):
        bib = {}
        bib['row'] = row
        bib['bibID'] = sheet1.cell(row, bibID_col_index).value
        bib['newOCN'] = sheet1.cell(row, newOCN_col_index).value
        # port['histcirc'] = sheet1.cell(row, histcirc_col_index).value
        # print(row)
        work_queue.put(bib)
    
    work_queue.join()
    output_queue.join()
    
    et = time.localtime()
    end_time = time.strftime("%H:%M:%S", et)
    print('Start Time: ', start_time)
    print('End Time: ', end_time)   

if __name__ == '__main__':
    main(sys.argv[1])
    