# -*- coding: utf-8 -*-
"""
Created on Thu Sep 14 09:15:46 2023
Applicants - Invventors network
"""
import pandas as pd
import matplotlib.pyplot as plt
from pathlib import Path


# Read Excel file
file_location = "C:\\Users\\Werner\\Desktop\\"
df = pd.read_excel(r"%sBiotech_comps.xlsx" %file_location)

# Group by 'Applicant Country' and 'Applicant' to count the number of patents for each applicant in each country
applicant_counts = df.groupby(['Applicant Country', 'Applicant'])['Title'].count().reset_index()

# Sort the DataFrame by the number of patents in descending order
applicant_counts = applicant_counts.sort_values(by='Title', ascending=False)

# Create an Excel writer to export data to multiple sheets
writer = pd.ExcelWriter('%sTop100-country.xlsx' %file_location, engine='openpyxl')

# Iterate through each country and write the top 100 applicants to separate sheets
for country in applicant_counts['Applicant Country'].unique():
    top_applicants = applicant_counts[applicant_counts['Applicant Country'] == country].head(100)
    sheet_name = f'Top 100 - {country}'
    top_applicants.to_excel(writer, sheet_name=sheet_name, index=False)

# Save the Excel file
writer._save()

