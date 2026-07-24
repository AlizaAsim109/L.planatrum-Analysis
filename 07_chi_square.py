# ==============================
# Chi-square Test for Gene Distribution
# ==============================

# Import required libraries
import numpy as np
from scipy.stats import chi2_contingency

# --------------------------------
# STEP 1: Define your data
# --------------------------------
# Replace these numbers with YOUR actual counts

# Format:
# [[Conserved_neuro, Conserved_non_neuro],
#  [Accessory_neuro, Accessory_non_neuro]]

contingency_table = np.array([
    [31, 2359],   # Conserved (core + soft core)
    [34, 9924]    # Accessory (shell + cloud)
])

# --------------------------------
# STEP 2: Run Chi-square test
# --------------------------------
chi2, p_value, dof, expected = chi2_contingency(contingency_table)

# --------------------------------
# STEP 3: Print results clearly
# --------------------------------
print("=== Chi-square Test Results ===\n")

print("Contingency Table:")
print(contingency_table, "\n")

print(f"Chi-square statistic: {chi2:.4f}")
print(f"Degrees of freedom: {dof}")
print(f"P-value: {p_value:.6f}\n")

print("Expected Frequencies:")
print(expected)

# --------------------------------
# STEP 4: Interpretation
# --------------------------------
alpha = 0.05

print("\n=== Interpretation ===")

if p_value < alpha:
    print("Result: Statistically significant difference detected (p < 0.05).")
    print("Interpretation: Gene distribution is NOT random across categories.")
else:
    print("Result: No statistically significant difference (p >= 0.05).")
    print("Interpretation: Gene distribution appears balanced across categories.")