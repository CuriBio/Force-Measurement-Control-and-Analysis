import openpyxl as xl
import os
import numpy as np
import matplotlib.pyplot as plt
from scipy.stats import linregress

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

mean_col = ws['D'] #um
stdev_col = ws['E'] #um
force_col = ws['G'] #uN

outputs = {
    "mean_disp_values": [],
    "stdev_disp_values": [],
    "force_values": []
}

for timepoint in range(2, len(mean_col)):
    outputs["mean_disp_values"].append(mean_col[timepoint].value)
    outputs["stdev_disp_values"].append(stdev_col[timepoint].value)
    outputs["force_values"].append(force_col[timepoint].value)

for output in outputs.keys():
    outputs[output] = np.array(outputs[output])
    outputs[output] = outputs[output].reshape((len(outputs[output])//7, 7))[:, :5]
    outputs[output] = outputs[output].astype(float)

# for low_post in range(7, 11):
#     plt.errorbar(outputs["force_values"][low_post, :],
#                  outputs["mean_disp_values"][low_post, :],
#                  outputs["stdev_disp_values"][low_post, :],
#                  marker='s',
#                  ms=5,
#                  capsize=10,
#                  color=colors[low_post],
#                  fmt='_')
#
#     result = linregress(outputs["force_values"][low_post, :], outputs["mean_disp_values"][low_post, :])
#     print(" R Sq.")
#     print(result.rvalue ** 2)
#     print(" Stiffness ")
#     print(result.slope**-1)
#     plt.plot(outputs["force_values"][low_post,:],
#              result.slope*outputs["force_values"][low_post, :] + result.intercept, color=colors[low_post])

def plot_f2p(post_range, color):
    mean_force = np.mean(outputs["force_values"][post_range, :], axis=0)
    mean_disp = np.mean(outputs["mean_disp_values"][post_range, :], axis=0)
    stdev_disp = np.std(outputs["mean_disp_values"][post_range, :], axis=0)

    plt.errorbar(mean_force,
                 mean_disp,
                 stdev_disp,
                 marker='s',
                 ms=5,
                 capsize=10,
                 color=color,
                 fmt='_')

    result = linregress(mean_force, mean_disp)
    print(" R Sq.")
    print(result.rvalue**2)
    print(" Stiffness ")
    print(result.slope**-1)
    plt.plot(mean_force,
             result.slope*mean_force + result.intercept, color=color)
#
# plot_f2p(slice(0, 6), "blue")
plot_f2p(slice(7, 11), "orange")
# plt.legend(["Low Stiffness Post", "High Stiffness Post"])
plt.title("high stiffness post")
plt.ylabel("post head displacement (\u03BCm)")
plt.xlabel("applied force (\u03BCN)")
plt.grid(True)
plt.savefig(r"G:\Shared drives\IR&D Projects\Mantarray-V1\3_Technical Development\Publication_Data" + "\\" + "high_stiffness_post")

plt.show()


print(output)