import random
import numpy as np
import xlrd
import xlwt
from xlwt import Workbook
# Workbook is created for writing
wb1 = Workbook()



loc = ("C:/Users/ADMIN/Desktop/python tools/microanalysis_pythonfull/micro_code.xlsx");
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
ns=sheet.nrows;#no of students
ni=sheet.ncols;#internal exams
###Question Paper details
total_internal_marks=50;
prt_a_tm=5;#part a question total marks
prt_a_qn=5;#part a question numbers
prt_a_qa=5;#part a question should attend
prt_a_eqm= prt_a_tm /prt_a_qa;#part a each question marks
if prt_a_qa!=prt_a_qn:
    prt_a_c=abs(prt_a_qa-prt_a_qn);#part b choice
else:
     prt_a_c=0
print("part a each q marks:",prt_a_eqm,"choice:", prt_a_c)
prt_b_tm=15;#part a question total marks
prt_b_qn=3;#part a question numbers
prt_b_qa=3;#part a question should attend
prt_b_eqm= prt_b_tm /prt_b_qa;#part a each question marks
if prt_b_qa!=prt_b_qn:
    prt_b_c=abs(prt_b_qa-prt_b_qn);#part b choice
else:
     prt_b_c=0
print("part b each q marks:",prt_b_eqm,"choice:", prt_b_c)

prt_c_tm=30;#part a question total marks
prt_c_qn=4;#part a question numbers
prt_c_qa=3;#part a question should attend
prt_c_eqm= prt_c_tm /prt_c_qa;#part a each question marks

if prt_c_qa!=prt_c_qn:
    prt_c_c=abs(prt_c_qa-prt_c_qn);#part b choice
else:
     prt_c_c=0

print("part c each q marks:",prt_c_eqm,"choice:", prt_c_c)


###Marks splitup functions
#tq=total questions, qa = question attended, eqm=each question mark, tm= total marks, ch=choice, student total mark in part a
def marksplit(tq,qa,eqm,stm,ch):
    l = [i for i in np.arange(0,eqm+0.5,0.5)]      
    r_prt_a_matrix = np.zeros(tq, dtype = float)
    r_tm=sum(r_prt_a_matrix)
    while r_tm!=stm:
        for i in range(tq):
            r_prt_a_matrix[i]=random.choice(l);
            if ch !=0:
                a=np.arange(1, tq+1, dtype=int)
                b=np.random.choice(a,ch,replace=False)
                for j in range(len(b)):
                    r_prt_a_matrix[b[j]-1]=0;              
        r_tm=sum(r_prt_a_matrix)
            
            
    
    return r_prt_a_matrix

### Student Section

#p = [i for i in np.arange(0,50.5,0.5)]
print(ni,ns)


for nii in range(1,ni+1):
    print(nii)
    print("Internal ", nii)
    if nii == 1:
        sheet1 = wb1.add_sheet('Sheet 1',cell_overwrite_ok=True);
    elif nii == 2:
        sheet1 = wb1.add_sheet('Sheet 2',cell_overwrite_ok=True)
    else:
        sheet1 = wb1.add_sheet('Sheet 3',cell_overwrite_ok=True)
    
    for nsi in range(1,ns+1):
        print(nsi)
        # For row 0 and column 0
        print(nsi,nii)
        print(sheet.cell_value(nsi-1, nii-1))
        student_marks=sheet.cell_value(nsi-1, nii-1);
        prt_a_stm=0;
        prt_b_stm=0;
        prt_c_stm=0;
        part_sum=prt_a_stm+prt_b_stm+prt_c_stm
        def ary(prt_fulmarks):
            g = [i for i in np.arange(0,prt_fulmarks+0.5,0.5)]
            return(random.choice(g))
            
        while part_sum!=student_marks:
            prt_a_stm=ary(prt_a_tm);
            prt_b_stm=ary(prt_b_tm);
            prt_c_stm=ary(prt_c_tm)
            part_sum=prt_a_stm+prt_b_stm+prt_c_stm

        print('students splitup marks: ', prt_a_stm,prt_b_stm,prt_c_stm)

        

        prt_a_matrix_f=marksplit(prt_a_qn,prt_a_qa,prt_a_eqm,prt_a_stm,prt_a_c)
        prt_b_matrix_f=marksplit(prt_b_qn,prt_b_qa,prt_b_eqm,prt_b_stm,prt_b_c)
        prt_c_matrix_f=marksplit(prt_c_qn,prt_c_qa,prt_c_eqm,prt_c_stm,prt_c_c)
        stu_splitup=np.concatenate([prt_a_matrix_f,prt_b_matrix_f,prt_c_matrix_f],axis=None)
        print(stu_splitup[1])
        print("nii is ",nii)
        sheet1.write(nsi-1, 0,stu_splitup[0])
        sheet1.write(nsi-1, 1,stu_splitup[1])
        sheet1.write(nsi-1, 2,stu_splitup[2])
        sheet1.write(nsi-1, 3,stu_splitup[3])
        sheet1.write(nsi-1, 4,stu_splitup[4])
        sheet1.write(nsi-1, 5,stu_splitup[5])
        sheet1.write(nsi-1, 6,stu_splitup[6])
        sheet1.write(nsi-1, 7,stu_splitup[7])
        sheet1.write(nsi-1, 8,stu_splitup[8])
        sheet1.write(nsi-1, 9,stu_splitup[9])
        sheet1.write(nsi-1, 10,stu_splitup[10])
        sheet1.write(nsi-1, 11,stu_splitup[11])
        wb1.save('xlwt example.xls')
        f_stu_marks=sum(stu_splitup)
        print(stu_splitup,f_stu_marks)
    
