import pandas as pd

# Specify the path to the Excel file
excel_path = "data.xlsx"

# Read the Excel file into a DataFrame
df = pd.read_excel(excel_path)


# Function to handle special characters
def clean_special_characters(text):
    # Remove special characters using regular expressions
    cleaned_text = ''.join(e for e in text if e.isalnum() or e.isspace())
    return cleaned_text



# Apply the clean_special_characters function to each cell in the DataFrame
for column in df.columns:
    df[column] = df[column].astype(str).apply(clean_special_characters)



# Function to add spaces before capitals
def add_space_before_capitals(text):
    # Initialize an empty string to store the modified text
    modified_text = ""
    # Iterate through each character in the text
    for i, char in enumerate(text):
        # If the character is a capital letter and the previous character is not a space, add a space before it
        if char.isupper() and i > 0 and text[i - 1] != ' ':
            modified_text += " " + char
        else:
            modified_text += char
    return modified_text.strip()  # Strip leading and trailing spaces


# Apply the add_space_before_capitals function to 'Product Name' column
df['Product Name'] = df['Product Name'].apply(add_space_before_capitals)
df['Supplier'] = df['Supplier'].apply(add_space_before_capitals)
df['Category'] = df['Category'].apply(add_space_before_capitals)


# Convert letters to uppercase
df['Supplier'] = df['Supplier'].str.title()
df['Category'] = df['Category'].str.title()
df['Product Name'] = df['Product Name'].str.title()



# Sort the DataFrame by Product ID
df['Product ID'] = df['Product ID'].astype(int)  # Convert Product ID to integer for sorting
df = df.sort_values(by='Product ID').reset_index(drop=True)

# Remove duplicate rows
df.drop_duplicates(inplace=True)

# Display the cleaned DataFrame
print(df)


# Specify the path for the cleaned Excel file
cleaned_excel_path = "cleaned_data.xlsx"

# Export the cleaned DataFrame to a new Excel file
df.to_excel(cleaned_excel_path, index=False)

print("Cleaned data has been exported to", cleaned_excel_path)