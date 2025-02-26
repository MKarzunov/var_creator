import os
import pandas as pd

basic_text = '''0
up
p_inE+06
p_outE+06
14
22
1
2.018
0.010
0.032
0.352
5.701E-003
1.9903E-002
1.01E-003
2.0E-003
1.0E-003
2.0E-003
1.01E-003
12
1
0.01355
7.50000E-004
7.50000E-004
1.4
287.1
T_in
7.14
6.4
1.00000E-03
1.00000E-03
0.010
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

    dir_name = str(row['t'])
    os.mkdir(dir_name)

    with open(os.path.join(dir_name, 'var.dat'), 'w') as var_file:
        var_file.write(new_text)

