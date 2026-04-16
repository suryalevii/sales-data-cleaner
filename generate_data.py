import pandas as pd
import numpy as np

data = {
    "OrderID":  [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17],
    "Customer": ["Ravi", "Meena", "Surya", "Arun", "Priya", "Karthik", "Tamil",
                 "Ravi", "Meena", "Surya", "Arun", "Priya", "Karthik", "Tamil",
                 "Ravi", "Meena", "Surya"],
    "Product":  ["Rice", "Oil", "Dal", "Sugar", "Tea", "Salt", "Rice",
                 "Oil", "Dal", "Sugar", "Tea", "Salt", "Rice", "Oil",
                 "Dal", "Sugar", "Tea"],
    "Category": ["Grain", "Oil", "Protein", "Sweet", "Drink", "Spice", "Grain",
                 "Oil", "Protein", "Sweet", "Drink", "Spice", "Grain", "Oil",
                 "Protein", "Sweet", "Drink"],
    "Quantity": [10, 5, 8, 3, 12, 7, 15, 4, 9, 6, 11, 2, 14, 8, 5, 10, 3],
    "Price":    [50, 120, 90, 40, 30, 20, 50, 120, 90, 40, 30, 20, 50, 120,
                 90, 40, 30],
    "Date":     ["2024-06-01", "2024-06-05", "2024-06-10", "2024-06-15",
                 "2024-06-20", "2024-07-01", "2024-07-05", "2024-07-10",
                 "2024-07-15", "2024-07-20", "2024-07-25", "2024-08-01",
                 "2024-08-05", "2024-08-10", "2024-08-15", "2024-08-20",
                 "2024-08-25"],
    "City":     ["Chennai", "Madurai", "Coimbatore", "Chennai", "Madurai",
                 "Coimbatore", "Chennai", "Madurai", "Coimbatore", "Chennai",
                 "Madurai", "Coimbatore", "Chennai", "Madurai", "Coimbatore",
                 "Chennai", "Madurai"]
}

df = pd.DataFrame(data)

# Add 3 duplicate rows
duplicates = df.iloc[[0, 5, 10]]
df = pd.concat([df, duplicates], ignore_index=True)

# Inject nulls
df.loc[2, "Price"] = None
df.loc[7, "Price"] = None
df.loc[3, "Product"] = None
df.loc[11, "Product"] = None

# Inject zero quantities
df.loc[4, "Quantity"] = 0
df.loc[9, "Quantity"] = 0

# Mix date formats
df.loc[1, "Date"] = "05/06/2024"
df.loc[6, "Date"] = "05/07/2024"

# Shuffle and reset
df = df.sample(frac=1, random_state=42).reset_index(drop=True)

# Save
df.to_excel("./Sales_project/raw_sales.xlsx", index=False)

print("Messy raw file created!")
print(f"Total rows: {len(df)}")
print("\nNulls per column:")
print(df.isnull().sum())
print(f"\nDuplicates: {df.duplicated().sum()}")