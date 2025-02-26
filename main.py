import os
import pandas as pd
import shutil
import subprocess

basic_text = '''0
up
p_inE+06
p_outE+06
16
30
1
2.127
0.023
0.032
0.366
4.000000E-03
4.737000E-03
1.000000E-03
4.200000E-03
3.200000E-03
4.200000E-03
1.000000E-03
12
1
0.013893
0.75000E-03
0.75000E-03
1.4
287.1
T_in
5.77
8.0
1.00000E-03
1.00000E-03
0.023
a
p
0.00001
0.00001
0.00002
0.00001
0.00'''

data = pd.read_excel("ГУ для расчета.xlsx", header=1)

for i, row in data.iterrows():
    if row['n'] == 0:
        up_limit = 1
    elif row['n'] < 1000:
        up_limit = 1000
    else:
        up_limit = 17000

    new_text = (basic_text
                .replace('up', str(up_limit))
                .replace('p_in', str(row['pвх*']))
                .replace('p_out', str(row['pвых']))
                .replace('T_in', str(row['Tвх*'])))

    dir_name = os.path.join('var', str(row['t']))
    try:
        os.mkdir(dir_name)
    except FileExistsError:
        pass

    with open(os.path.join(dir_name, 'var.dat'), 'w') as var_file:
        var_file.write(new_text)
        shutil.copy("app_with_sas.exe", dir_name)
        # subprocess.run(os.path.join(dir_name, "app_with_sas.exe"))

