import fitz 
import re
import pandas as pd

final_list=[]

#a dictionary to maintain a list of Programme name in the PDF versus the Programme name format I want in my Database.
# This is only for TY SFC. We need to manually feed this list according to the PDF to be extracted.
#When doing this for Other Programmes dont replace the current dictionary.
#Comment out the current dictionary used and assign the new desired dictionay. This way next time it becomes easier to fetch the details of previous programmes again.
course_map={"TYBVoc":"TYBVOC","TYBAMMC":"TYBAMMC","TYBCOM-AF":"TYBCOM-AF","TYBCOM-BI":"TYBCOM-BI","TYBMS":"TYBMS","TYBSC-B.T.":"TYBSC-BT","TYBSC-IT":"TYBSC-IT"}

def extract_data_from_pdf(pdf_path):
    doc = fitz.open(pdf_path)

    #This list contains the names according to the names specified in the PDF. Important to match it exactly with the programme name in the PDF
    courses_list=["TYBVoc","TYBAMMC","TYBCOM-AF","TYBCOM-BI","TYBMS","TYBSC-B.T.","TYBSC-IT"]
    #pattern_for_name=re.compile(r'^[A-Z][A-Z]{1,}$')
    pattern_for_name=re.compile(r'^[A-Z][\'\"\'\']{0,2}[A-Z]{1,}$')
    
    for page_num in range(doc.page_count):
        page = doc[page_num]
        print(f"\nPAGE IS {page}")
        page_text = page.get_text()
       
        page_data=[]
        control_id=[]
        names=[]
        current_name=[]
        current_course=""

        main_iterator=iter(page_text.split("\n"))
        one_step_ahead_of_main_iterator=iter(page_text.split("\n"))

        next(one_step_ahead_of_main_iterator,None)
        
        for row, next_row_value in zip(main_iterator,one_step_ahead_of_main_iterator):
            #print(f"EACH ROW IS:{row.split()}")
            
            #destructure each row
            current_row= row.split()
            #print("NEW ROW:",current_row)
            
            if(len(current_row)==1):
                [current_word]=current_row
                if(current_word.startswith("20") and len(current_word)==10):
                    control_id.append(current_word)
                    
                elif(pattern_for_name.match(current_word) and current_word not in courses_list ):
                    #print("name1 is",current_word)
                    current_name.append(current_word)
                elif(current_word in courses_list):
                    current_course=course_map[current_word]
                    #print(f"Current course:{current_course}")
                    

            elif(len(current_row)==2):
                [name1, name2]=current_row
                #print(name1,name2)
                if(pattern_for_name.match(name1) and name1 not in courses_list and pattern_for_name.match(name2) and name2 not in courses_list):
                    #print(f" NAME1 is:{name1} and NAME2 is :{name2}")
                    current_name.extend([name1,name2])
                    
            elif(len(current_row)==3):
                [name1, name2,name3]=current_row
                if(pattern_for_name.match(name1) and name1 not in courses_list and pattern_for_name.match(name2) and name2 not in courses_list and pattern_for_name.match(name3) and name3 not in courses_list):
                    #print(f" NAME1 is:{name1} and NAME2 is :{name2} and NAME3 is :{name3}")
                    current_name.extend([name1,name2,name3])
            
            #to check whether the name is done by checking whether the next column is that of the subjects/courses
            #if the column is of courses, it contains "-"(hypen) and ","(comma)
                    
            if("," in next_row_value or "-" in next_row_value):
                #print(f"next row is:{next_row_value}, Count:{len(next_row_value)}")
                if(len(current_name)>0):
                    temp_list=[]
                    temp_list.append(current_course)
                    temp_list.extend(current_name)
                    #current_name.append(current_course)
                    #names.append(current_name)
                    names.append(temp_list)
                current_name=[]
                
                    
            

            
        print(f"Control id: {control_id} {len(control_id)}")
        
        print(f"Names: {names}{len(names)}")
        create_final_list(control_id,names,final_list)
        
    doc.close()
    print(f"\n...................... FINAL LIST(Total: {len(final_list)})................\n")
    print(final_list)


def create_final_list(control_id_list,names_list,final_list):
    for control_id, name in zip(control_id_list,names_list):
        temp_list=[]
        temp_list.append(control_id)
        temp_list.extend(name)
        final_list.append(temp_list)
        
        
def export_to_excel_using_pandas(final_list,excel_file_path):
    df=pd.DataFrame(final_list,columns=["Control ID","Course","Last Name","First Name","Middle Name","Mothers Name","",""])
    try:
        df.to_excel(excel_file_path,index=False)
        print("FILE EXPORTED TO EXCEL")
    except Exception as e:
        print("Error occurred while writing to Excel")
        print(e)
    
#Manually set the path of the file where we need to extract the data
#THE PATH SHOULD CONTAIN FORWARD SLASHES INSTEAD OF BACKSLASHES ELSE IT TREATS IT LIKE AN ESCAPE CHARACTER   
extract_data_from_pdf("D:/OTHER PROJECTS/Extract_Student_info/PDFS/SY-All-Course.pdf")

export_to_excel_using_pandas(final_list,"D:/OTHER PROJECTS/Extract_Student_info/EXCEL_SHEETS/Student_info_SY_SFC.xlsx")
