import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import pandas as pd

driver=webdriver.Chrome()

link= "https://charusat.edu.in:912/UniExamResult/frmUniversityResult.aspx"


driver.get(link)


def get_student_data(student_id):
    
    institute_select = Select(driver.find_element(By.ID, 'ddlInst'))
    institute_select.select_by_visible_text('CSPIT')
    
    
    degree_select = Select(driver.find_element(By.ID, 'ddlDegree'))
    degree_select.select_by_visible_text('BTECH (AIML)')
    
    
    semester_select = Select(driver.find_element(By.ID, 'ddlSem'))
    semester_select.select_by_visible_text('3')
    
    
    exam_select = Select(driver.find_element(By.ID, 'ddlScheduleExam'))
    exam_select.select_by_visible_text('NOVEMBER 2024')
    
    
    student_id_input = driver.find_element(By.ID, 'txtEnrNo')
    student_id_input.clear()
    student_id_input.send_keys(student_id)
    
   
    show_marksheet_button = driver.find_element(By.ID, 'btnSearch')
    show_marksheet_button.click()
        
    
    time.sleep(2)
    
    try:
        student_name_element = driver.find_element(By.ID, 'uclGrd1_lblStudentName') 
        student_name = student_name_element.text.strip() 
        student_sgpa_element = driver.find_element(By.ID, 'uclGrd1_lblSGPA')
        student_sgpa = student_sgpa_element.text.strip()
        back_button = driver.find_element(By.ID, 'ibtnBack')
        back_button.click() 
    except Exception as e: 
        print(f"Failed to retrieve data for {student_id}: {e}") 
        student_name = None 
    return student_name, student_id, student_sgpa


students_data = []

for student_id in range(1,83):
    student_id_str = f"23AIML{student_id:03d}"
    try:
        student_name, student_id, student_sgpa = get_student_data(student_id_str)
        if student_name:
            students_data.append([student_name, student_id, student_sgpa])
            print(f"Retrieved data for {student_id_str}: {student_name}")
    except Exception as e:
        print(f"Failed to retrieve data for {student_id_str}: {e}")


driver.quit() 


detail = pd.DataFrame(students_data, columns=['Student Name', 'Student ID', 'SGPA']) 

detail.to_excel('students_data.xlsx', index=False) 

print("Data has been successfully saved to students_data.xlsx")




