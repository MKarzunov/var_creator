import os
import shutil
import pandas as pd
from scipy.optimize import curve_fit
from matplotlib import pyplot as plt


def linear(x, k, b):
    return k * x + b


def poly3(x, a3, a2, a1, a0):
    return a3 * (x ** 3) + a2 * (x ** 2) + a1 * x + a0


data = pd.read_excel("Для аппроксимации.xlsx", header=None, index_col=None)
revolutions = data.iloc[list(range(2, data.shape[0], 7)), 3:]
torque = data.iloc[list(range(5, data.shape[0], 7)), 3:]
time = data.iloc[list(range(2, data.shape[0], 7)), 0]
triplets = []
print(f"Data sets: {revolutions.shape[0]}")
use_poly = {
    0.05: False,
    0.1: False,
    0.2: False,
    0.3: False,
    0.4: False,
    0.5: False,
    0.6: False,
    0.7: False,
    0.79: False}

if os.path.isdir('time'):
    shutil.rmtree('time')
os.mkdir('time')

linear_formulas = ''

for i in range(revolutions.shape[0]):
    triplet = [revolutions.iloc[i], torque.iloc[i], time.iloc[i]]

    linear_params = curve_fit(linear, triplet[0], triplet[1])[0]
    poly3_params = curve_fit(poly3, triplet[0], triplet[1])[0]

    upper_bound = int(triplet[0].iloc[-1] * 1.05) + 1
    x_values = list(range(upper_bound))

    linear_values = [linear(item, *linear_params) for item in x_values]
    poly3_values = [poly3(item, *poly3_params) for item in x_values]

    fig, ax = plt.subplots()
    ax.grid()
    ax.set_title(f"time {triplet[2]}s")
    ax.set_xlim(0, upper_bound)
    ax.set_xlabel("RPM")
    ax.set_ylabel("Torque, Nm")
    ax.plot(triplet[0], triplet[1], marker="o", linestyle="")
    ax.plot(x_values, linear_values, linestyle="dashed")
    ax.plot(x_values, poly3_values, linestyle="dashed")
    ax.legend(['Initial data', 'Linear approximation', 'Polynomial approximation'])
    fig.savefig(f"time\\time {round(triplet[2], 2)}s.png")
    plt.close()

    triplet.append(linear_params)
    triplet.append(poly3_params)

    triplets.append(triplet)
    linear_formulas += f't = {time.iloc[i]} | M = {linear_params[0]} * n + {linear_params[1]}\n'

with open("linear_formulas.txt", 'w') as formulas_file:
    formulas_file.write(linear_formulas)

# if os.path.isdir('rotation'):
#     shutil.rmtree('rotation')
# os.mkdir('rotation')

# n_values = list(range(200, 1000, 200))
# n_values.extend(list(range(1000, 18001, 1000)))
#
# for n in n_values:
#     time = []
#     torque = []
#     for triplet in triplets:
#         current_time = triplet[2]
#         time.append(current_time)
#         if use_poly[round(current_time, 2)]:
#             torque.append(poly3(n, *triplet[4]))
#         else:
#             torque.append(linear(n, *triplet[3]))
#
#     fig, ax = plt.subplots()
#     ax.grid()
#     ax.set_title(f"{n} RPM")
#     ax.set_xlabel("time, s")
#     ax.set_ylabel("Torque, Nm")
#     ax.plot(time, torque)
#     fig.savefig(f"rotation\\{n} RPM.png")
#     plt.close()

