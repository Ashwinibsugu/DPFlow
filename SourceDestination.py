import pandas as pd

# Read the LBG sheet
first_sheet_df = pd.read_excel('lbglatest.xlsx')

# Read the TCP sheet
second_sheet_df = pd.read_excel('fshlatest - Copy.xlsx')

# Initialize a list to store Domain values
domain_values = []

# Iterate through each row of the LBG sheet
for _, first_row in first_sheet_df.iterrows():
    target_ip = first_row['Port']
    #sever_ip = first_row['Server']

    # Initialize a variable to store the Domain value
    domain = None

    # Iterate through each row of the TCP to find the matching 'targetIP'
    for _, second_row in second_sheet_df.iterrows():
        #print("Checking:", target_ip, second_row['targetIP'])
        if int(second_row['targetIP']) == int(target_ip):
            domain = second_row['Domain']
            break  # Stop iterating once a match is found

    domain_values.append(domain)

# Add the Domain values to the first sheet
first_sheet_df['Destination'] = domain_values

# Remove rows where 'Destination' is empty
#df = first_sheet_df.dropna(subset=['Destination'])

# Save the updated dataframe back to the first sheet
first_sheet_df.to_excel('first_sheet_with_domain.xlsx', index=False)

third_sheet_df = pd.read_excel('first_sheet_with_domain.xlsx')
condition1 = third_sheet_df['Health'] != 'down'
#condition2 = third_sheet_df['Domain'] != 'RBSDBD_HO'
#condition3 = third_sheet_df['Domain'] != 'RBSMPD_EX'
#condition4 = third_sheet_df['Domain'] != 'RBSMBK_HO4'

# Apply the filtering
filtered_df = third_sheet_df[condition1]
#filtered_df = third_sheet_df[condition2]
#filtered_df = third_sheet_df[condition3]
#filtered_df = third_sheet_df[condition4]

fourth_sheet_df = filtered_df[~filtered_df['Port'].isin([443, 444])]

# Optionally, you can reset the index of the filtered DataFrame
fourth_sheet_df.reset_index(drop=True, inplace=True)

# Filter out rows where 'Destination' is empty
filtered_df1 = fourth_sheet_df[fourth_sheet_df['Destination'].notna()]

# Define conditions to filter rows
condition1 = filtered_df1['LBG Group'].str.startswith('RBS')  # Rows where 'LBG Group' starts with 'RBS'
condition2 = ~filtered_df1['LBG Group'].str.contains('internal|zcee|inbound|json|canary|edb|mps', case=False, regex=True)  # Rows where 'LBG Object' does not contain 'internal' or 'zcee'
#condition3 = filtered_df['Domain'].isin(['RBSDBD_HO', 'RBSMPD_EX', 'RBSMBK_HO4'])
# Apply both conditions to filter the DataFrame
filtered_df5 = filtered_df1[condition1 & condition2]

#condition4 = filtered_df5['Domain'] != filtered_df['Destination']
#filtered_df6 = filtered_df5[condition4]


# Save the filtered DataFrame to a new Excel file
filtered_df5.to_excel('FinalSourceDestination.xlsx', index=False)