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

ResultPath = r'C:\\Users\\Danil\\Downloads\\Telegram Desktop\\'  # Путь до файла проекта и для сохранения TOUCHSTONE

# Topology_1 - полоса заграждения выше по частоте, Topology_2 - ниже по частоте, Topology_3 - симметрично;

File = 'Topology_3.cst'  # Имя проекта

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
        writer = csv.writer(csvfile, quoting=csv.QUOTE_NONE, escapechar=' ')

        if not file_exists:
            writer.writerow(['Name project', 'Current datetime', 'Filter topology',
                             'Frequency start (Mhz)', 'Frequency end (Mhz)',
                             'Frequency start_2 (Mhz)', 'Frequency end_2 (Mhz)',
                             'Frequency start_3 (Mhz)', 'Frequency end_3 (Mhz)',
                             'attenuation level in the passband (dB)', 'attenuation level in the stopband (dB)',
                             'attenuation level in the stopband_2 (dB)', 'H (mm)', 'HH (mm)',
                             'L1 (mm)', 'L11 (mm)', 'L2 (mm)', 'L3 (mm)',
                             'Lp (mm)', 'Ls (mm)', 'Rgnd (mm)', 'Rp (mm)',
                             'S1_1 (mm)', 'S2_1 (mm)', 'S3_1 (mm)', 'T (mm)',
                             'W1_1 (mm)', 'W2_1 (mm)', 'W3_1 (mm)', 'Wp (mm)',
                             'The path to the Touchstone', 'The path to the file'])
        rounded_list = map(lambda x: round(x, 3), data_list)
        list1 = map(str, rounded_list)
        writer.writerow([name_project, datetime.now(), name_project.split('.')[0], a, b, a_2, b_2, a_3, b_3, db_1, db_2, db_3, ','.join(list1), name_2, name_1])

'''
def change_par(mycst1, name, values):
    par_opt_1_1 = 'Optimizer.SelectParameter ("' + str(name) + '", True)' \
                  '\n Optimizer.SetParameterMin (' + str(values) + ')' \
                  '\nEnd Sub'
    mycst1.schematic.execute_vba_code(par_opt_1_1, timeout=None)


def apply_change_par(mycst1, df, par_name, MIN):
    mask = (df['Name'].isin(par_name)) & (df['Values'] < MIN)
    for name in df.loc[mask, 'Name']:
        print(name)
        change_par(mycst1, MIN, name)'''

def optim(mycst1, a_1, b_1, db_1, a_2, b_2, db_2, delta, a_3='-', b_3='-', db_3='-'):
    resultFile = cst.results.ProjectFile(ResultPath + str(File), allow_interactive=True)
    schematic = resultFile.get_schematic()
    r = pd.DataFrame(data=list(schematic.get_parameter_combination(0).items()), columns=['Name', 'Values'])
    print('Checking for the minimum value of physical parameters')
    #apply_change_par(mycst1, r, ['L1', 'L2', 'L3'], L_min)

    L1 = L_min if (r[r['Name']=='L1']['Values']).any() < L_min else r.loc[r['Name']=='L1', 'Values'].iloc[0]
    L2 = L_min if (r[r['Name']=='L2']['Values']).any() < L_min else r.loc[r['Name']=='L2', 'Values'].iloc[0]
    L3 = L_min if (r[r['Name']=='L3']['Values']).any() < L_min else r.loc[r['Name']=='L3', 'Values'].iloc[0]
    L11 = W_min if (r[r['Name']=='L11']['Values']).any() < W_min else r.loc[r['Name']=='L11', 'Values'].iloc[0]
    S1_1 = S_min if (r[r['Name']=='S1_1']['Values']).any() < S_min else r.loc[r['Name']=='S1_1', 'Values'].iloc[0]
    S2_1 = S_min if (r[r['Name']=='S2_1']['Values']).any() < S_min else r.loc[r['Name']=='S2_1', 'Values'].iloc[0]
    S3_1 = S_min if (r[r['Name']=='S3_1']['Values']).any() < S_min else r.loc[r['Name']=='S3_1', 'Values'].iloc[0]
    W1_1 = W_min if (r[r['Name']=='W1_1']['Values']).any() < W_min else r.loc[r['Name']=='W1_1', 'Values'].iloc[0]
    W2_1 = W_min if (r[r['Name']=='W2_1']['Values']).any() < W_min else r.loc[r['Name']=='W2_1', 'Values'].iloc[0]
    W3_1 = W_min if (r[r['Name']=='W3_1']['Values']).any() < W_min else r.loc[r['Name']=='W3_1', 'Values'].iloc[0]

    par_opt_1 = 'Sub Main () \n Optimizer.DeleteAllGoals \n Dim goalID As Long \n goalID = Optimizer.AddGoal("1DC Primary Result")' \
                '\n  Optimizer.SetAndUpdateMinMaxAuto (' + str(delta) + ')' \
                '\n  Optimizer.SelectParameter ("L1", True)' \
                '\n  Optimizer.SetParameterMin (' + str(L1) + ')' \
                '\n  Optimizer.SelectParameter ("L2", True)' \
                '\n  Optimizer.SetParameterMin (' + str(L2) + ')' \
                '\n  Optimizer.SelectParameter ("L3", True)' \
                '\n  Optimizer.SetParameterMin (' + str(L3) + ')' \
                '\n  Optimizer.SelectParameter ("L11", True)' \
                '\n  Optimizer.SetParameterMin (' + str(L11) + ')' \
                '\n  Optimizer.SelectParameter ("S1_1", True)' \
                '\n  Optimizer.SetParameterMin (' + str(S1_1) + ')' \
                '\n  Optimizer.SelectParameter ("S2_1", True)' \
                '\n  Optimizer.SetParameterMin (' + str(S2_1) + ')' \
                '\n  Optimizer.SelectParameter ("S3_1", True)' \
                '\n  Optimizer.SetParameterMin (' + str(S3_1) + ')' \
                '\n  Optimizer.SelectParameter ("W1_1", True)' \
                '\n  Optimizer.SetParameterMin (' + str(W1_1) + ')' \
                '\n  Optimizer.SelectParameter ("W2_1", True)' \
                '\n  Optimizer.SetParameterMin (' + str(W2_1) + ')' \
                '\n  Optimizer.SelectParameter ("W3_1", True)' \
                '\n  Optimizer.SetParameterMin (' + str(W3_1) + ')' \
                '\n  Optimizer.SetAndUpdateMinMaxAuto (' + str(delta) + ')' \
                 '\n  Optimizer.SelectGoal (goalID, True)' \
                 '\n  Optimizer.SetGoal1DCResultName(".\S-Parameters\S1,1")' \
                 '\n  Optimizer.SetGoalScalarType("magdB20")' \
                 '\n  Optimizer.SetGoalOperator ("<")' \
                 '\n  Optimizer.SetGoalTarget (' + str(db_1) + ')' \
                 '\n  Optimizer.SetGoalWeight (4.0)' \
                 '\n  Optimizer.SetGoalRangeType ("range")' \
                 '\n  Optimizer.SetGoalRange (' + str(a_1) + ', ' + str(b_1) + ')' \
                 '\nEnd Sub'
    mycst1.schematic.execute_vba_code(par_opt_1, timeout=None)
    print('The purpose has been updated')
    par_opt_end = 'Sub Main () \n Dim goalID As Long \n goalID = Optimizer.AddGoal("1DC Primary Result")' \
                '\n  Optimizer.SelectGoal (goalID, True)' \
                '\n  Optimizer.SetGoal1DCResultName(".\S-Parameters\S2,1")' \
                '\n  Optimizer.SetGoalScalarType("magdB20")' \
                '\n  Optimizer.SetGoalOperator ("<")' \
                '\n  Optimizer.SetGoalTarget (' + str(db_2) + ')' \
                '\n  Optimizer.SetGoalWeight (1.0)' \
                '\n  Optimizer.SetGoalRangeType ("range")' \
                '\n  Optimizer.SetGoalRange (' + str(a_2) + ', ' + str(b_2) + ')' \
                '\nEnd Sub'

    mycst1.schematic.execute_vba_code(par_opt_end, timeout=None)
    if topology == 3:
        par_opt_3 = 'Sub Main () \n Dim goalID As Long \n goalID = Optimizer.AddGoal("1DC Primary Result")' \
                    '\n  Optimizer.SelectGoal (goalID, True)' \
                    '\n  Optimizer.SetGoal1DCResultName(".\S-Parameters\S2,1")' \
                    '\n  Optimizer.SetGoalScalarType("magdB20")' \
                    '\n  Optimizer.SetGoalOperator ("<")' \
                    '\n  Optimizer.SetGoalTarget (' + str(db_3) + ')' \
                    '\n  Optimizer.SetGoalWeight (1.0)' \
                    '\n  Optimizer.SetGoalRangeType ("range")' \
                    '\n  Optimizer.SetGoalRange (' + str(a_3) + ', ' + str(b_3) + ')' \
                    '\nEnd Sub'
        mycst1.schematic.execute_vba_code(par_opt_3, timeout=None)

    mycst1.model3d.run_solver()

    par_opt_start = 'Sub Main ()' \
                    '\n  Optimizer.Start' \
                    '\nEnd Sub'

    print('Optimization')
    #mycst1.schematic.execute_vba_code(par_opt_start, timeout=None)
    resultFile = cst.results.ProjectFile(ResultPath + str(File), allow_interactive=True)

    res_filename = File + '_' + str(a_1) + '_' + str(b_1)

    par_opt_impr = 'Sub Main ()' \
                   '\n TOUCHSTONE.Reset' \
                   '\n TOUCHSTONE.FileName("' + ResultPath + res_filename + '_S_param' + '")' \
                   '\n TOUCHSTONE.ExportType ("S")' \
                   '\n TOUCHSTONE.Format ("MA")' \
                   '\n TOUCHSTONE.FrequencyRange("Full")' \
                   '\n TOUCHSTONE.Write' \
                   '\n End Sub'
    print('saving to a file')
    mycst1.schematic.execute_vba_code(par_opt_impr, timeout=None)

    # Parameters schematic------------------------------------------------------------------------------
    res_fileparam = 'param_' + res_filename
    schematic = resultFile.get_schematic()
    schematic.get_all_run_ids()

    write_to_csv(File[:-6] + '.csv', File, str(a_1), str(b_1), str(db_1), str(a_2), str(b_2), str(db_2), str(a_3),str(b_3), str(db_3), list(schematic.get_parameter_combination(0).values()),ResultPath + res_filename + '_S_param', ResultPath + File)
    print("Update CSV file")
    return 'good'


def Solve(A, B, F1, F2, F3, Ch, Ch_2, Step, db_1, db_2, db_3):
    update = 17  # скорость обновления ограничений физ параметров
    mycst = cst.interface.DesignEnvironment()
    mycst1 = cst.interface.DesignEnvironment.open_project(mycst, ResultPath + str(File))
    while (B - A):
        if (B - A) < 0:
            break
        a_1 = A - A * F1 / 100 * 0.5
        b_1 = A + A * F1 / 100 * 0.5
        a_2 = A - A * F2 / 100 * 0.5 + A * Ch / 100
        b_2 = A + A * F2 / 100 * 0.5 + A * Ch / 100
        delta = update * (A + Step) / A  # Расчет процента изменения параметров
        if topology == 3:
            a_3 = A - A * F3 / 100 * 0.5 - A * Ch_2 / 100
            b_3 = A + A * F3 / 100 * 0.5 - A * Ch_2 / 100
            optim(mycst1, a_1, b_1, db_1, a_2, b_2, db_2, delta, a_3, b_3, db_3)
            # print(a_1, b_1, db_1, a_2, b_2, db_2, a_3, b_3, db_3, update, delta)
            print('--------------------------------------------------')

        else:
            optim(mycst1, a_1, b_1, db_1, a_2, b_2, db_2, delta)
            # print(a_1, b_1, db_1, a_2, b_2, db_2, update, delta)
            print('--------------------------------------------------')

        A += Step
        update += 1
    cst.interface.DesignEnvironment.close(mycst)


A = float(input('Начальная центральная частота (МГц): '))
B = float(input('Конечная частота(МГц): '))
F1 = float(input('Процент полосы пропускания (10): '))
F2 = float(input('Процент полосы заграждения (10): '))
Ch = float(input('Процент отступа от центральной частоты до верхней полосы заграждения (20): '))
Step = float(input('Шаг смещения центральной частоты (МГц): '))
db_1 = float(input('Уровень согласования в полосе пропускания (дБ): '))
db_2 = float(input('Уровень затухания в верхней полосе заграждения (дБ): '))
F3, Ch_2, db_3 = '-', '-', '-'
if topology == 3:
    F3 = float(input('Процент полосы заграждения два (10): '))
    Ch_2 = float(input('Процент отступа от центральной частоты до нижней полосы заграждения (20): '))
    db_3 = float(input('Уровень затухания в нижней полосе заграждения (дБ): '))
L_min = float(input('Нижнее ограничение физических параметров - длин фильтра (L1, L2, L3): '))
W_min = float(input('Нижнее ограничение физических параметров - ширины фильтра (W1, W2, W3, L11): '))
S_min = float(input('Нижнее ограничение физических параметров - расстояние между связнными линиями фильтра (S1_1, S2_1, S3_1): '))

Solve(A, B, F1, F2, F3, Ch, Ch_2, Step, db_1, db_2, db_3)