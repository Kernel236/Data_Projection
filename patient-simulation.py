#------------SETUP-------------
import os
import pandas as pd
import numpy as np
import scipy.stats as stats
import matplotlib.pyplot as plt
import seaborn as sns

# Excel file path
file_path = "data/Report_Dose_Tracking_Rieti_Maggio_2023.xlsx"

# Read the Excel file
excel_file = pd.ExcelFile(file_path)

# Check the names of the sheets in the Excel file
# .SHEET_NAMES is a property of the ExcelFile object that returns a list of all sheet names in the file.
sheet_names = excel_file.sheet_names

#Initialize an empty dict to store DataFrames
dataframes_simulated = {}

# Loop through each sheet name and read the data into a DataFrame
for sheet in sheet_names:
    print(f"\n----- Reading sheet: {sheet} -----")
    # Read the sheet into a DataFrame
    df = pd.read_excel(file_path, sheet_name=sheet)

    #Verify the normality of the data in the numerical columns (Età, Dose_Erogata)
    def normality_test(v, name, sheet_name):
        fig = plt.figure()
        stats.probplot(v, dist="norm", plot=plt)
        plt.title(f"QQ Plot Test for {name} ({sheet_name})")
        plt_filename = f"QQ_Plot_Test_{sheet_name}_{name}.png"
        plt.savefig(plt_filename)
        plt.show()
        plt.close(fig)

        #More formal test for normality
        stat, p = stats.shapiro(v)
        print(f"Shapiro-Wilk test for {name}: W={stat} p-value={p}")
        return p > 0.05
    
    età_normal = normality_test(df['Età'], 'Età', sheet)
    dose_normal = normality_test(df['Dose_Erogata'], 'Dose_Erogata', sheet)

    #Simulate the data
    # Create a new DataFrame with the same structure as the original
    if età_normal:
        età_simulated = np.random.normal(loc=df['Età'].mean(), scale=df['Età'].std(), size=120)
    else:
        età_simulated = np.random.choice(df['Età'], size=120, replace=True)

    # Simulate the Dose_Erogata column
    # Check if the data is normally distributed or not
    if dose_normal:
        dose_simulated = np.random.normal(loc=df['Dose_Erogata'].mean(), scale=df['Dose_Erogata'].std(), size=120)
    else:
        dose_simulated = np.random.choice(df['Dose_Erogata'], size=120, replace=True)

    # Build projection for Categorical data
    #p = df['Sesso'].value_counts(normalize=True).values is the probability distribution of the categorical variable 'Sesso'
    # Simulate the Sesso column
    sesso_categories = df['Sesso'].unique()
    sesso_probabilities = df['Sesso'].value_counts(normalize=True).reindex(sesso_categories, fill_value=0).values
    sex_simulated = np.random.choice(sesso_categories, size=120, replace=True, p=sesso_probabilities)

    #Simulate mounth 
    mounth_categories = df['Mese'].unique()
    mounth_probabilities = df['Mese'].value_counts(normalize=True).reindex(mounth_categories, fill_value=0).values
    mounth_simulated = np.random.choice(mounth_categories, size=120, replace=True, p=mounth_probabilities)

    #Build the new DataFrame with the simulated data
    simulated_df = pd.DataFrame({
        'Età': np.round(età_simulated),
        'Sesso': sex_simulated,
        'Dose_Erogata': np.round(dose_simulated, 2),
        'Mese': mounth_simulated
    })

    dataframes_simulated[sheet] = simulated_df

# Write Excel file with multiple sheets
output_path = "Output/Report_Dose_Tracking_Rieti_Maggio_2023_120_pazienti.xlsx"
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    for sheet_name, df in dataframes_simulated.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f"Simulated data for {sheet_name} written to Excel file.")

#Confirmation message
print("All sheets have been processed and saved to the new Excel file in Output.")