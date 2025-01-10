import os
import sys
sys.path.insert(0, os.path.abspath('C:\\Program Files (x86)\\CST Studio Suite 2024\\AMD64\\python_cst_libraries'))  # Поменять путь к библиотеке CST
import cst
import cst.interface
from cst.interface import Project
import cst.results
import numpy as np

mycst = cst.interface.DesignEnvironment()
mycst1 = cst.interface.DesignEnvironment.open_project(mycst, r'C:\\Users\\Danil\\Downloads\\Telegram Desktop\\Start_5.cst')  # Поменять путь к проекту CST

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
project = cst.results.ProjectFile(r'C:\\Users\\Danil\\Downloads\\Telegram Desktop\\Start_5.cst')
S11 = project.get_3d().get_result_item(r"1D Results\S-Parameters\S1,1")
res = np.array(S11.get_ydata())
print(res)




