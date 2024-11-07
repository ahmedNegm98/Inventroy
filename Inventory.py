from bs4 import BeautifulSoup
import pandas as pd
import os
import chardet

def extract_hardware_info(html_content, pc_location):
    # Parse the HTML content
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # Dictionary to store hardware info with lists for repeated fields
    hardware_info = {
        'PC Location': [pc_location],
        'Computer Name': [],
        'Computer Brand Name': [],
        'Product Serial Number': [],
        'Motherboard Model': [],
        'CPU Brand Name': [],
        'Total Memory Size': [],
        'Video Card': [],
        'Monitor Name (Manuf)': [],
        'Media Rotation Rate': [],
        'Drive Capacity': []
    }

    # Extract fields for all keys in hardware_info
    for label in hardware_info.keys():
        sections = soup.find_all(string=label + ":")
        for section in sections:
            value = section.find_next().text.strip().splitlines()[0]
            # Handle 'Computer Brand Name' specifically for duplicates
            if label == 'Computer Brand Name':
                if value not in hardware_info[label]:  # Check for duplicates
                    hardware_info[label].append(value)
            else:
                if hardware_info[label]:
                    hardware_info[label][0] += "\n" + value
                else:
                    hardware_info[label].append(value)

    return hardware_info

# Directory containing the HTML files
directory_path = 'F:/AndCreativeAds/inventoryDevices/windows'
data = []

for filename in os.listdir(directory_path):
    if filename.endswith('.HTM') or filename.endswith('.html'):
        html_file_path = os.path.join(directory_path, filename)
        pc_location = os.path.splitext(filename)[0]

        # Read the raw bytes first for encoding detection
        with open(html_file_path, 'rb') as file:
            raw_data = file.read()

        # Detect encoding
        result = chardet.detect(raw_data)
        encoding = result['encoding']

        try:
            # Decode using the detected encoding
            html_content = raw_data.decode(encoding)

            # Extract hardware info
            hardware_info = extract_hardware_info(html_content, pc_location)

            # Prepare the data for each extracted hardware info
            max_length = max(len(hardware_info['Media Rotation Rate']), len(hardware_info['Drive Capacity']))
            
            for i in range(max_length):
                row = {
                    'Media Rotation Rate': hardware_info['Media Rotation Rate'][i] if i < len(hardware_info['Media Rotation Rate']) else None,
                    'Drive Capacity': hardware_info['Drive Capacity'][i] if i < len(hardware_info['Drive Capacity']) else None,
                    'PC Location': hardware_info['PC Location'][0]
                }
                data.append(row)

            for key in hardware_info.keys():
                if key not in ['Media Rotation Rate', 'Drive Capacity', 'PC Location']:
                    for row in data[-max_length:]:
                        row[key] = hardware_info[key][0] if hardware_info[key] else None

        except UnicodeDecodeError:
            print(f"Could not read {html_file_path}. Skipping this file.")

# Convert to DataFrame
df = pd.DataFrame(data)
column_order = ['PC Location'] + list(hardware_info.keys())[1:]
df = df[column_order]

# Save to Excel file
excel_path = 'F:/Python/Inventory/hardware_info.xlsx'
df.to_excel(excel_path, index=False)

print(f"Hardware information has been saved to {excel_path}.")
