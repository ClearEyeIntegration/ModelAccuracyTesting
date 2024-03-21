import time
import supportingdoc_extraction
from util import get_access_token, Excel_Path, UPLOAD_WAIT

# IF more files are uploaded, please increase the wait time
#UPLOAD_WAIT = 200 #amitabh commented-time in seconds
#UPLOAD_WAIT = 600



def run_test():
    # try:
    
    response = get_access_token()
    if response.status_code == 200:
        token = response.json()['result']['access_token']
        print(token)
        reference_id_list = supportingdoc_extraction.upload_files(token)
        print('Upload completed and reference list is ' + str(reference_id_list))
        #print('waiting for upload to complete')
        #time.sleep(UPLOAD_WAIT)
        if len(reference_id_list) != 0:
            supportingdoc_extraction.list_of_jobs(token, reference_id_list)
            supportingdoc_extraction.remove_duplicate_entries()
            supportingdoc_extraction.remove_expectedsheet_entries()
            supportingdoc_extraction.compare_fields(Excel_Path)
            supportingdoc_extraction.accuracy(Excel_Path)
        else:
            print("Job's list are empty")
            exit()
    else:
        print('Issue with Authentication and the response code is ' + str(response.status_code))
        print(response)
        exit()


run_test()
