import os
import sys
sys.path.insert(0, os.path.abspath('C:\\Program Files (x86)\\CST Studio Suite 2024\\AMD64\\python_cst_libraries'))  # Поменять путь к библиотеке CST
import cst
import cst.interface
from cst.interface import Project
import cst.results
import numpy as np
import pandas as pd
import csv
from datetime import datetime

File = 'Topology_3.cst'   #Имя проекта


topology = int(File.split('.')[0].split('_')[1])

''' #Изменение Parameter list------------------------------------------------------------------------------
par_change = 'Sub Main () \nStoreParameter("W3_1", 2.4)\nStoreParameter("S3_1",2)' \
             '\nRebuildOnParametricChange (bfullRebuild, bShowErrorMsgBox)' \
             '\nStoreParameter("L3", 3)' \
             '\nEnd Sub'
mycst1.schematic.execute_vba_code(par_change, timeout=None)
'''
'''
def save_S_P(s1, s2, s3, s4, s5, s6, s7, s8, name_f):
    s = [s1, s2, s3, s4, s5, s6, s7, s8]
    try:
        with open(name_f, 'w', encoding='utf-8') as file:
            for s_1 in s:
                mm = ','.join(map(str, s_1))
                file.write(mm + '\n\n\n\n\n\n\n')
        print(f"S-Parameters have been successfully saved to a file: {name_f}")
    except Exception as e:
        print(f"An error occurred when writing to a file: {e}") '''

def write_to_csv(filename, name_project, a, b, db_1, a_2, b_2, db_2, a_3, b_3, db_3, data_list, name_2, name_1):
    file_exists = False
    try:
        with open(filename, 'r'):
            file_exists = True
    except FileNotFoundError:
        pass
    with open(filename, 'a', newline='') as csvfile:
        writer = csv.writer(csvfile,quoting=csv.QUOTE_NONE, escapechar=' ')

        if not file_exists:
            writer.writerow(['Name project', 'Current datetime', 'Filter topology',
                             'Frequency start (Ghz)', 'Frequency end (Ghz)',
                             'Frequency start_2 (Ghz)', 'Frequency end_2 (Ghz)',
                             'Frequency start_3 (Ghz)', 'Frequency end_3 (Ghz)',
                             'attenuation level in the passband (dB)', 'attenuation level in the stopband (dB)',
                             'attenuation level in the stopband_2 (dB)', 'H (millimetre)' , 'HH (millimetre)' ,
                             'L1 (millimetre)', 'L11 (millimetre)', 'L2 (millimetre)', 'L3 (millimetre)',
                             'Lp (millimetre)', 'Ls (millimetre)', 'Rgnd (millimetre)', 'Rp (millimetre)',
                             'S1_1 (millimetre)', 'S2_1 (millimetre)', 'S3_1 (millimetre)', 'T (millimetre)',
                             'W1_1 (millimetre)', 'W2_1 (millimetre)', 'W3_1 (millimetre)', 'Wp (millimetre)',
                             'The path to the Touchstone', 'The path to the file'])
        rounded_list = map(lambda x: round(x, 3), data_list)
        list1 = map(str, rounded_list)
        writer.writerow([name_project, datetime.now(),name_project.split('.')[0], a, b, a_2, b_2, a_3, b_3, db_1, db_2, db_3, ','.join(list1), name_2, name_1])

def optim(mycst1, a_1, b_1, db_1, a_2, b_2, db_2, i, a_3 = '-', b_3 = '-', db_3 = '-'):
    par_opt_1 = 'Sub Main () \n Optimizer.DeleteAllGoals \n Dim goalID As Long \n goalID = Optimizer.AddGoal("1DC Primary Result")' \
                '\n  Optimizer.SelectParameter ("L1", True)' \
                '\n  Optimizer.SelectParameter ("L2", True)' \
                '\n  Optimizer.SelectParameter ("L3", True)' \
                '\n  Optimizer.SelectParameter ("L11", True)' \
                '\n  Optimizer.SelectParameter ("S1_1", True)' \
                '\n  Optimizer.SelectParameter ("S2_1", True)' \
                '\n  Optimizer.SelectParameter ("S3_1", True)' \
                '\n  Optimizer.SelectParameter ("W1_1", True)' \
                '\n  Optimizer.SelectParameter ("W2_1", True)' \
                '\n  Optimizer.SelectParameter ("W3_1", True)' \
                '\n  Optimizer.SelectGoal (goalID, True)' \
                '\n  Optimizer.SetGoal1DCResultName(".\S-Parameters\S1,1")' \
                '\n  Optimizer.SetGoalTarget (' + str(db_1) + ')' \
                '\n  Optimizer.SetGoalWeight (1.0)' \
                '\n  Optimizer.SetGoalRangeType ("range")' \
                '\n  Optimizer.SetGoalRange (' + str(a_1) + ', ' + str(b_1) + ')' \
                '\n  Optimizer.SetAndUpdateMinMaxAuto (' + str(i) + ')' \
                '\nEnd Sub'
    par_opt_2 = 'Sub Main () \n Dim goalID As Long \n goalID = Optimizer.AddGoal("1DC Primary Result")' \
                '\n  Optimizer.SelectGoal (goalID, True)' \
                '\n  Optimizer.SetGoal1DCResultName(".\S-Parameters\S2,1")' \
                '\n  Optimizer.SetGoalTarget (' + str(db_2) + ')' \
                '\n  Optimizer.SetGoalWeight (1.0)' \
                '\n  Optimizer.SetGoalRangeType ("range")' \
                '\n  Optimizer.SetGoalRange (' + str(a_2) + ', ' + str(b_2) + ')' \
                '\nEnd Sub'

    mycst1.schematic.execute_vba_code(par_opt_1, timeout=None)
    mycst1.schematic.execute_vba_code(par_opt_2, timeout=None)
    if topology == 3:
        par_opt_3 = 'Sub Main () \n Dim goalID As Long \n goalID = Optimizer.AddGoal("1DC Primary Result")' \
        '\n  Optimizer.SelectGoal (goalID, True)' \
        '\n  Optimizer.SetGoal1DCResultName(".\S-Parameters\S2,1")' \
        '\n  Optimizer.SetGoalTarget (' + str(db_3) + ')' \
        '\n  Optimizer.SetGoalWeight (1.0)' \
        '\n  Optimizer.SetGoalRangeType ("range")' \
        '\n  Optimizer.SetGoalRange (' + str(a_3) + ', ' + str(b_3) + ')' \
        '\nEnd Sub'
        mycst1.schematic.execute_vba_code(par_opt_3, timeout=None)

    mycst1.modeler.run_solver()

    par_opt_start = 'Sub Main ()' \
                    '\n  Optimizer.Start' \
                    '\nEnd Sub'
    mycst1.schematic.execute_vba_code(par_opt_start, timeout=None)
    resultFile = cst.results.ProjectFile(r'C:\\Users\\Danil\\Downloads\\Telegram Desktop\\' +str(File), allow_interactive=True)

    #S-Parameters schematic------------------------------------------------------------------------------
    S11 = resultFile.get_3d().get_result_item(r"1D Results\S-Parameters\S1,1")
    S12 = resultFile.get_3d().get_result_item(r"1D Results\S-Parameters\S1,2")
    S21 = resultFile.get_3d().get_result_item(r"1D Results\S-Parameters\S2,1")
    S22 = resultFile.get_3d().get_result_item(r"1D Results\S-Parameters\S2,2")

    S11_y = S11.get_ydata()
    S11_x = S11.get_xdata()
    S12_y = S12.get_ydata()
    S12_x = S12.get_xdata()
    S21_y = S21.get_ydata()
    S21_x = S21.get_xdata()
    S22_y = S22.get_ydata()
    S22_x = S22.get_xdata()
    res_filename = File + '_' + str(a_1) + '_' + str(b_1)

    #txt_headers = ['S11_y', 'S11_x', 'S12_y', 'S12_x', 'S21_y', 'S21_x', 'S22_y', 'S22_x']
    ###save_S_P(S11_y, S11_x, S12_y, S12_x, S21_y, S21_x, S22_y, S22_x, res_filename)
    par_opt_impr = 'Sub Main ()' \
    '\n TOUCHSTONE.Reset' \
    '\n TOUCHSTONE.FileName("'+ res_filename +'_S_param'+'")' \
    '\n TOUCHSTONE.ExportType ("S")' \
    '\n TOUCHSTONE.Format ("MA")' \
    '\n TOUCHSTONE.FrequencyRange("Full")'\
    '\n TOUCHSTONE.Write' \
    '\n End Sub'
    mycst1.schematic.execute_vba_code(par_opt_impr, timeout=None)

   #Parameters schematic------------------------------------------------------------------------------
    res_fileparam = 'param_' + res_filename
    schematic = resultFile.get_schematic()
    schematic.get_all_run_ids()
    r = pd.DataFrame(data = list(schematic.get_parameter_combination(0).items()), columns = ['Name','Values'])
    r.set_index('Name', inplace=True)

    write_to_csv(File[:-6] + '.csv', File, str(a_1) , str(b_1), str(db_1), str(a_2), str(b_2), str(db_2), str(a_3), str(b_3), str(db_3), list(schematic.get_parameter_combination(0).values()), res_filename +'_S_param', os.path.abspath(File))
    print("Update CVS file")
    return 'good'

def Solve(A, B, F1, F2, F3, Ch, Ch_2, Step, db_1, db_2, db_3):
    update = 10  # скорость обновления ограничений физ параметров
    mycst = cst.interface.DesignEnvironment()
    mycst1 = cst.interface.DesignEnvironment.open_project(mycst, r'C:\\Users\\Danil\\Downloads\\Telegram Desktop\\' + str(File))  # Поменять путь к проекту CST
    while (B - A):
        if (B - A) < 0:
            break
        a_1 = A - A * F1/100 * 0.5
        b_1 = A + A * F1/100 * 0.5
        a_2 = A - A * F2 / 100 * 0.5 + A * Ch/100
        b_2 = A + A * F2 / 100 * 0.5 + A * Ch/100
        if topology == 3:
            a_3 = A - A * F3 / 100 * 0.5 - A * Ch_2/100
            b_3 = A + A * F2 / 100 * 0.5 - A * Ch / 100
            optim(mycst1, a_1, b_1, db_1, a_2, b_2, db_2, update, a_3, b_3, db_3)
            #print(a_1, b_1, db_1, a_2, b_2, db_2, a_3, b_3, db_3, update)
            print('--------------------------------------------------')

        else:
            optim(mycst1, a_1, b_1, db_1, a_2, b_2, db_2, update)
            #print(a_1, b_1, db_1, a_2, b_2, db_2, update)
            print('----------------------------------')

        A += Step
        update+=1
    cst.interface.DesignEnvironment.close(mycst)

A = float(input('Целевая частота: '))
B = float(input('Конечная частота: '))
F1 = float(input('Процент полосы пропускания (10): '))
F2 = float(input('Процент полосы заграждения (10): '))
Ch = float(input('Отступ от целевой (12): '))
Step = float(input('Шаг: '))
db_1 = float(input('db_1: '))
db_2 = float(input('db_2: '))
F3, Ch_2, db_3 = '-', '-', '-'
if topology == 3:
    F3 = float(input('Процент полосы заграждения два (10): '))
    Ch_2 = float(input('Отступ от целевой два (12): '))
    db_3 = float(input('db_3: '))

Solve(A, B, F1, F2, F3, Ch, Ch_2, Step, db_1, db_2, db_3)


'''
a1 = float(input('Values start : '))
b1 = float(input('Values end : '))
c1 = float(input('Change : '))
if topology == 3:
    c2 = float(input('-Change_2 : '))
step = float(input('Step : '))
wind = float(input('Range : '))
db_1 = float(input('Values_1 dB: '))
db_2 = float(input('Values_2 dB: '))
if topology == 3:
    db_3 = float(input('Values_3 dB: '))

#solv(MIN Первой цели, MAX Первой цели, отступ Второй цели, ШАГ,  Размер окна)
def solv(a1, b1, c1, step_1, wind, c2=0):
    j = 0
    if c2:
        for i in np.arange(a1, b1, step_1):
            min_1 = round(i,3) + step_1*j
            max_1 = round(i,3) + wind + step_1*j
            min_2 = round(i,3) + wind + c1 + step_1*j
            max_2 = round(i,3) + wind*2 + c1 + step_1*j
            min_3 = round(i,3) - wind - c2 + step_1*j
            max_3 = round(i,3) - c2 + step_1*j
            j+=1
            optim(min_1, max_1, min_2, max_2, min_3, max_3)
    else:
        for i in np.arange(a1, b1, step_1):
            min_1 = round(i,3) + step_1*j
            max_1 = round(i,3) + wind + step_1*j
            min_2 = round(i,3) + wind + c1 + step_1*j
            max_2 = round(i,3) + wind*2 +c1 + step_1*j
            j += 1
            print(min_1, max_1, min_2, max_2)
            optim(min_1, max_1, min_2, max_2)

if topology == 3:
    solv(a1, b1, c1, step, wind,c2)
else:
    solv(a1, b1, c1, step, wind)
'''


