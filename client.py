# Import necessary libraries
import pandas as pd
import xlsxwriter

# Read data from HDB dataset and selection criteria
dataFrame = pd.read_excel('all.xlsx', sheet_name='Table 1')  # Read HDB dataset
selections = pd.read_excel('selections.xlsx', sheet_name='Sheet1')  # Read selection criteria

# Group data by building for further analysis
dff = dataFrame.groupby(['Building'])['Cell', 'Model', 'Building', 'Apartment', 'Floor', 'Area', 'Privilege', 'Total Price']

# Extract relevant selection criteria
buildings = selections['Selected Building'].values.tolist()  # List of selected buildings
baseApartments = selections['Reference Apartment'].values.tolist()  # List of reference apartments
priorities = selections['Priority'].values.tolist()  # List of priorities

# Filter the dataset based on specific criteria
categoryBuilding = dataFrame.loc[(dataFrame['Building'].isin(buildings)) & (dataFrame['Area'] > 115) & (dataFrame['Floor'] != 5)]

# Define a function to determine if an apartment corresponds to the reference apartment on the sea wind direction for each building
def correspondsTo(apartmentNum, base):
    if ((apartmentNum - base) % 4) == 0:
        return True
    else:
        return False

print(len(baseApartments))

# Filter and color the apartments based on priority
data = []
selectionsIterator = 0

# Iterate over selected buildings
for building in buildings:
    dt = categoryBuilding.loc[(categoryBuilding['Building'] == buildings[selectionsIterator]) & ((categoryBuilding['Apartment'] - baseApartments[selectionsIterator]) % 4 == 0)]
    dt['Priority'] = dt['Priority'].replace(0, priorities[selectionsIterator])
    data.append(dt)
    selectionsIterator += 1

# Combine the filtered data from all buildings
allSelections = pd.concat(data, ignore_index=True, sort=True)

# Define a function to highlight and color-code the rows based on priority
def highlight_rows(row):
    pr = row.loc['Priority']
    if pr == 1:
        color = '#FFB3BA'  # Green
    elif pr == 2:
        color = '#FFA500'  # Grey
    else:
        color = '#FCFCFA'
    return ['background-color: {}'.format(color) for r in row]

# Select and sort columns for final presentation
allSelections = allSelections[['Total Price', 'Privilege', 'Model', 'Priority', 'Floor', 'Apartment', 'Building', 'Cell']]
allSelections = allSelections.sort_values(by=['Cell'], ascending=True)

# Apply the conditional formatting to the data
dt = allSelections.style.apply(highlight_rows, axis=1)

# Save the formatted data to an Excel file
dt.to_excel("best.xlsx", index=False)
