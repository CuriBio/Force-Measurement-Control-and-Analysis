import openpyxl as xl
import os
import numpy as np
import matplotlib.pyplot as plt
from scipy.stats import linregress
from scipy.stats import shapiro

spreadsheet_path = os.path.join(r"G:\Shared drives\IR&D Projects\Mantarray-V1\3_Technical Development\MAv1_Field-to-Force\Position To Force\Spring Constant Estimates -- Hanging Weight\spring_constant_test_11252022\weight_hanging_test_set_1.xlsx")
print(spreadsheet_path)

wb = xl.load_workbook(spreadsheet_path, data_only=True)
ws = wb.active

raw_col = np.array(ws['A'])
raw_disp_values_list = []
weight_num_12x = np.arange(6)
weight_12x = .000050*9.8 # kg * m/s sq
scale = 71000 # pixels / m
num_posts = 5
num_cycles = 4
num_weights = len(weight_num_12x)

force = weight_num_12x * weight_12x

for cell in range(0, len(raw_col)):
    if type(raw_col[cell].value) == float or type(raw_col[cell].value) == int:
        raw_disp_values_list.append(raw_col[cell].value)

# absolute pixel position of post head
position_values = np.array(raw_disp_values_list).reshape(num_posts, num_cycles, num_weights)/scale
# change in position from unweighted position
displacement_values = (position_values.T - position_values[:, :, 0].T).T
# mean change in unweighted position
mean_disp_values = np.mean(displacement_values, axis=1)
# standard deviation of change in unweighted position
stdev_disp_values = np.std(displacement_values, axis=1)

# Example position - to - force
for post in range(0, 5):

    ## plot mean displacements
    # plt.errorbar(
    #     force,
    #     mean_disp_values[post],
    #     stdev_disp_values[post],
    #     marker='s',
    #     ms=5,
    #     capsize=10,
    #     fmt='_'
    # )
    #
    # result = linregress(force, mean_disp_values[post, 0])

    # plot displacements from one loading cycle
    plt.plot(
        force,
        displacement_values[post, 0],
        "x"
    )

    result = linregress(force, displacement_values[post, 0])

    print(" R Sq.")
    print(result.rvalue**2)
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

# Plot all regression slopes
for post in range(0, 5):
    # result = linregress(force, mean_disp_values[post]) # mean displacement of all cycles
    result = linregress(force, displacement_values[post, 0]) # one cycle

    print(" R Sq.")
    print(result.rvalue**2)
    print(" Stiffness ")
    print(result.slope**-1)
    spring_constants_12x.append(result.slope**-1)

k_mean = np.mean(spring_constants_12x)
k_2stdev = 2*np.std(spring_constants_12x)
print("mean stiffness:")
print(k_mean)
print("2x standard deviation stiffness")
print(k_2stdev)
print("as a percentage of mean stiffness")
print(k_2stdev/k_mean * 100)

print("Shapiro-Wilk p-value")
stat, p = shapiro(spring_constants_12x)
print(p)

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
for k in range(0, len(spring_constants_12x)):
    plt.plot(bar_positions[0]
             + .1*np.random.rand()
             - .05,
             spring_constants_12x[k], marker="${0}$".format(k + 1), color="black")


plt.ylabel("stiffness (N/m)")
plt.grid(True)
plt.show()




