import openpyxl as xl
import os
import numpy as np
import matplotlib.pyplot as plt
from scipy.stats import linregress
from scipy.stats import shapiro

spreadsheet_path = os.path.join(r"C:\Users\kevin\OneDrive\Documents\weight_hanging_test_set_1.xlsx")
print(spreadsheet_path)

wb = xl.load_workbook(spreadsheet_path, data_only=True)
ws = wb.active

raw_col = np.array(ws['A'])
raw_disp_values_list = []
weight_num_12x = np.arange(6)
weight_12x = .000050 * 9.8 # kg * m/s sq
scale = 71000 # pixels / m

force = weight_num_12x * weight_12x

for cell in range(0, len(raw_col)):
    if type(raw_col[cell].value) == float or type(raw_col[cell].value) == int:
        raw_disp_values_list.append(raw_col[cell].value)

raw_disp_values = np.array(raw_disp_values_list).reshape(5, 4, 6) / scale
mean_disp_values = np.mean(raw_disp_values.T - raw_disp_values[:, :, 0].T, axis=1).T
stdev_disp_values = np.std(raw_disp_values.T - raw_disp_values[:, :, 0].T, axis=1).T

### Example position - to - force
for post in range(0, 5):
    plt.errorbar(force,
                 mean_disp_values[post],
                 stdev_disp_values[post],
                 marker='s',
                 ms=5,
                 capsize=10,
                 fmt='_')

    result = linregress(force, mean_disp_values[post])
    print(" R Sq.")
    print(result.rvalue ** 2)
    print(" Stiffness ")
    print(result.slope**-1)
    plt.plot(force,
             result.slope*force + result.intercept)

    plt.ylabel("post head displacement (\u03BCm)")
    plt.xlabel("applied force (\u03BCN)")
    plt.title("post {0}".format(post + 1))
    plt.grid(True)
    plt.tight_layout()
    plt.show()

spring_constants_12x = []

### Example position - to - force
for post in range(0, 5):
    result = linregress(force, mean_disp_values[post])
    print(" R Sq.")
    print(result.rvalue ** 2)
    print(" Stiffness ")
    print(result.slope ** -1)
    spring_constants_12x.append(result.slope**-1)

bar_positions = [1]
plt.margins(.125)
plt.bar(bar_positions,
             [np.mean(spring_constants_12x)],
             yerr=[2*np.std(spring_constants_12x)],
             capsize=10,
        tick_label=["high stiffness"],
        edgecolor="black",
        facecolor="gainsboro",
        color="white",
        width=.2)
plt.xticks(bar_positions)
plt.plot(bar_positions * np.ones(len(spring_constants_12x)) + .1 * np.random.rand(len(spring_constants_12x)) - .05
         , spring_constants_12x, "x", color="black")


plt.ylabel("stiffness (N/m)")
plt.grid(True)
plt.show()

print(np.mean(spring_constants_12x))

# Test Normality of each timepoint
stat, p = shapiro(spring_constants_12x)
print(p)


