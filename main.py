import os
import sys
sys.path.insert(0, os.path.abspath('C:\\Program Files (x86)\\CST Studio Suite 2024\\AMD64\\python_cst_libraries'))  # Поменять путь к библиотеке CST
import cst
import cst.interface
from cst.interface import Project
import cst.results
import numpy as np
import pandas as pd

File = 'Start_5.cst'   #Имя проекта

mycst = cst.interface.DesignEnvironment()
mycst1 = cst.interface.DesignEnvironment.open_project(mycst, r'C:\\Users\\Danil\\Downloads\\Telegram Desktop\\' + str(File))  # Поменять путь к проекту CST

''' #Изменение Parameter list
par_change = 'Sub Main () \nStoreParameter("W3_1", 2.4)\nStoreParameter("S3_1",2)' \
             '\nRebuildOnParametricChange (bfullRebuild, bShowErrorMsgBox)' \
             '\nStoreParameter("L3", 3)' \
             '\nEnd Sub'
mycst1.schematic.execute_vba_code(par_change, timeout=None)
'''

input_str = input('Values for the S-1,1 range (e.g., 1.5,2.5 or 1.5 2.5): ')
a_str, b_str = input_str.split(',') if ',' in input_str else input_str.split()
a_1 = float(a_str)
b_1 = float(b_str.strip())

input_str = input('Values for the S-2,1 range (e.g., 1.5,2.5 or 1.5 2.5): ')
a_str, b_str = input_str.split(',') if ',' in input_str else input_str.split()
a_2 = float(a_str)
b_2 = float(b_str.strip())

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
            '\n  Optimizer.SetGoalTarget (-20)' \
            '\n  Optimizer.SetGoalWeight (1.0)' \
            '\n  Optimizer.SetGoalRangeType ("range")' \
            '\n  Optimizer.SetGoalRange (' + str(a_1) + ', ' + str(b_1) + ')' \
                                                                          '\nEnd Sub'

par_opt_2 = 'Sub Main () \n Dim goalID As Long \n goalID = Optimizer.AddGoal("1DC Primary Result")' \
            '\n  Optimizer.SelectGoal (goalID, True)' \
            '\n  Optimizer.SetGoal1DCResultName(".\S-Parameters\S2,1")' \
            '\n  Optimizer.SetGoalTarget (-45)' \
            '\n  Optimizer.SetGoalWeight (1.0)' \
            '\n  Optimizer.SetGoalRangeType ("range")' \
            '\n  Optimizer.SetGoalRange (' + str(a_2) + ', ' + str(b_2) + ')' \
                                                                          '\nEnd Sub'

mycst1.schematic.execute_vba_code(par_opt_1, timeout=None)
mycst1.schematic.execute_vba_code(par_opt_2, timeout=None)

mycst1.modeler.run_solver()

par_opt_start = 'Sub Main ()' \
                '\n  Optimizer.Start' \
                '\nEnd Sub'
mycst1.schematic.execute_vba_code(par_opt_start, timeout=None)

cst.interface.DesignEnvironment.close(mycst)

mycst = cst.interface.DesignEnvironment()
resultFile = cst.results.ProjectFile(r'C:\\Users\\Danil\\Downloads\\Telegram Desktop\\' +str(File), allow_interactive=True)

#project = cst.results.ProjectFile( r'C:\\Users\\Danil\\Downloads\\Telegram Desktop\\Start_5.cst')
S11 = resultFile.get_3d().get_result_item(r"1D Results\S-Parameters\S1,1")
S12 = resultFile.get_3d().get_result_item(r"1D Results\S-Parameters\S1,2")
S21 = resultFile.get_3d().get_result_item(r"1D Results\S-Parameters\S2,1")
S22 = resultFile.get_3d().get_result_item(r"1D Results\S-Parameters\S2,2")

res_filename = File + '_' + str(a_1) + '_' + str(b_1)

csv_headers = ['S11_y', 'S11_x', 'S12_y', 'S12_x', 'S21_y', 'S21_x', 'S22_y', 'S22_x']



#Parameters schematic
res_fileparam = 'param_' + res_filename
schematic = resultFile.get_schematic()
schematic.get_all_run_ids()
r = pd.DataFrame(data = list(schematic.get_parameter_combination(0).items()), columns = ['Name','Values'])
r.set_index('Name', inplace=True)
r.to_csv(res_fileparam, mode='a', header=True, index=True, encoding='utf-8')



