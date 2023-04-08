import numpy as np
import matplotlib.pyplot as plt
import math

def calculate_k(rating):
    k_min = 16
    k_max = 32
    mid_point = 2500
    steepness = 0.004
    
    # Logistic function
    k = k_min + (k_max - k_min) / (1 + math.exp(-steepness * (rating - mid_point)))
    return k

ratings = np.arange(0, 5000, 10)
k_values = [calculate_k(rating) for rating in ratings]

plt.plot(ratings, k_values)
plt.xlabel('Player Rating')
plt.ylabel('K Value')
plt.title('K Value as a Function of Player Rating')
plt.grid()
plt.show()
