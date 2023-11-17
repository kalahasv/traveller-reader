import PyPDF2 as pf
from pathlib import Path


def convert_to_job(text):
    text_arr = text.splitlines()
    #how to know what belongs to what? can use position for the  first 6 if I split lines
    #what after that? for certificates and supplier qualifications i can look for keywords and use a boolean
    #it looks like tolerance is always there: look for text following "tolerance:". May or may not have locations after it. 
    # Seems like it's always two lines so i can just grab the next line
    info_line_1 = text_arr[2] #contains job id, date, and contact information
    #regex match to J_ID format



    print(text_arr)




def get_line(pdfFile):
    with open(pdfFile) as f:
        reader = pf.PdfReader(pdfFile)

        print(len(reader.pages))

        page = reader.pages[0]

        text = page.extract_text()
        text2 = page.get_contents()
        print(text)
        print("Conversion")
        convert_to_job(text)



if __name__ == '__main__':

    file_name = '057531C-traveler.pdf'
    path =  Path('jfiles',file_name)
    get_line(path)