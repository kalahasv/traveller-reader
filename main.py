import PyPDF2 as pf
from pathlib import Path
import pdftotree
from bs4 import BeautifulSoup
import pandas as pd
from IPython.display import display
import tabula
from format_traveler import create_excel

def pdf_to_df(file_name):
    df = tabula.read_pdf(file_name, pages='all')

if __name__ == '__main__':

    file_name = '05695AD-traveler.pdf'
    path =  Path('jfiles',file_name)
    pdf_to_df(path)
    create_excel()