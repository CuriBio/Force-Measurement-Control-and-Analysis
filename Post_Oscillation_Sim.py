import numpy as np
from scipy.integrate import solve_ivp
import sys

print(sys.path)
sys.path.insert(0, r'C:\Users\NSB\Documents\GitHub\Accelerated_Magnet_Finding')
print(sys.path)

from Accelerated_Magnet_Finding import Process_Data_UT

m_mag = .001 * (.00075/2)**2 * np.pi * 7.5 * 1000
m_post_head = 5e-6 - m_mag/7.5 + m_mag
m_post_neck = 1.1e-5
m = m_post_head + .23*m_post_neck
b = .0005
k = .159
w0 = np.sqrt(abs(b**2 - 4*m*k))/2/m
w0_hz = w0/2/np.pi
print(w0_hz)