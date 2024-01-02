from create_excel import create_excel
import process as p

if __name__ == '__main__':

   # file_name = '0593318-traveler.pdf' #OP - Deburr - Finish - Inserts - PM (Laser) - Final Inspection - Bag&Tag
    #file_name = '052BD56-traveler.pdf' #OP - Deburr - PM(Engraving) - Finish - Final Inspection - Bag&Tag
    #file_name = '057531C-traveler.pdf' # OP - Deburr - Final Inspection - Bag&Tag
    #file_name = '05695AD-traveler.pdf' # OP - Deburr - Finish - Final - Bag&Tag
    #file_name = '05AC673-traveler.pdf'
    p.init()
    create_excel(p.get_trav_df(),p.get_process_df(),p.get_dd())
    