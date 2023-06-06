import openpyxl as xl
import os
import numpy as np
import matplotlib.pyplot as plt
from scipy.stats import linregress
from scipy.stats import shapiro

spreadsheet_path = os.path.join("G:\Shared drives\IR&D Projects\Mantarray-V1\\3_Technical Development\Publication_Data",
                                "Hanging Spring Constant Measurements.xlsx")
print(spreadsheet_path)

wb = xl.load_workbook(spreadsheet_path, data_only=True)
ws = wb.active


colors = [
    "orange",
    "black",
    "blue",
    "red",
    "green",
    "grey"
]
colors = colors + colors

raw_col = np.array(ws['A:C'])
mean_col = ws['D'] #um
stdev_col = ws['E'] #um
force_col = ws['G'] #uN

raw_disp_values = []
mag_nums_1x = (np.arange(5) + 1)*2
mag_num_12x = np.arange(5) + 1
mag_weight_1x = 2.2 * 9.8 # mg * m/s sq
mag_weight_12x = 60.5 * 9.8 # mg * m/s sq

mag_force = []
mag_force.append(mag_nums_1x * mag_weight_1x)
mag_force.append(mag_num_12x * mag_weight_12x)

for timepoint in range(2, len(mean_col)):
    for cell in range(0, raw_col.shape[0]):
        if type(raw_col[cell, timepoint].value) == float or type(raw_col[cell, timepoint].value) == int:
            raw_disp_values.append(raw_col[cell, timepoint].value)

raw_disp_values = np.array(raw_disp_values).reshape(len(raw_disp_values)//15, 5, 3)
mean_disp_values = np.mean(raw_disp_values, axis=2)
stdev_disp_values = np.std(raw_disp_values, axis=2)

# ### Example position - to - force
# for post in [0, 7]:
#     plt.errorbar(mag_force[post//6],
#                  mean_disp_values[post],
#                  stdev_disp_values[post],
#                  marker='s',
#                  ms=5,
#                  capsize=10,
#                  color=colors[post],
#                  fmt='_')
#
#     result = linregress(mag_force[post//6], mean_disp_values[post])
#     print(" R Sq.")
#     print(result.rvalue ** 2)
#     print(" Stiffness ")
#     print(result.slope**-1)
#     plt.plot(mag_force[post//6],
#              result.slope*mag_force[post//6] + result.intercept, color=colors[post])
#
# plt.ylabel("post head displacement (\u03BCm)")
# plt.xlabel("applied force (\u03BCN)")
# plt.legend(["low stiffness", "high stiffness"])
# plt.grid(True)
# plt.savefig(r"G:\Shared drives\IR&D Projects\Mantarray-V1\3_Technical Development\Publication_Data" + "\\" + "post_stiffness_example")
# plt.show()



spring_constants_1x = []
spring_constants_12x = []

### Example position - to - force
for post in range(0, 11):
    result = linregress(mag_force[post // 6], mean_disp_values[post])
    print(" R Sq.")
    print(result.rvalue ** 2)
    print(" Stiffness ")
    print(result.slope ** -1)

    if (post < 6):
        spring_constants_1x.append(result.slope**-1)
    elif (post > 6):
        spring_constants_12x.append(result.slope**-1)

["low stiffness", "high stiffness"]
# plt.figure(figsize=(5, 6))
bar_positions = [.375, .625]
plt.margins(.125)
plt.bar(bar_positions,
             [np.mean(spring_constants_1x), np.mean(spring_constants_12x)],
             yerr=[2*np.std(spring_constants_1x), 2*np.std(spring_constants_12x)],
             capsize=10,
        tick_label=["low stiffness", "high stiffness"],
        edgecolor="black",
        facecolor="gainsboro",
        color="white",
        width=.2)
plt.xticks(bar_positions)
plt.plot(bar_positions[0] * np.ones(len(spring_constants_1x)) + .1 * np.random.rand(len(spring_constants_1x)) - .05,
         spring_constants_1x, "x", color="black")
plt.plot(bar_positions[1] * np.ones(len(spring_constants_12x)) + .1 * np.random.rand(len(spring_constants_12x)) - .05
         , spring_constants_12x, "x", color="black")


plt.ylabel("stiffness (N/m)")
plt.grid(True)
plt.savefig(r"G:\Shared drives\IR&D Projects\Mantarray-V1\3_Technical Development\Publication_Data" + "\\" + "spring constant variability")
plt.show()

print(spring_constants)

# Test Normality of each timepoint
stat, p = shapiro(spring_constants)
print(p)


