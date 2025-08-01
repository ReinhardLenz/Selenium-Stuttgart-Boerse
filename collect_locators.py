import re

# Path to the Python file to parse
file_path = r"c:\\temp\\Lib\\POMProjectFolder\\Stuttgart\\Test\\global_variable_sharing_example\\Bond_search6.py"

# Read the file
with open(file_path, 'r', encoding='utf-8') as f:
    code = f.read()

# Find all patterns like Locators.something)
pattern = re.findall(r'Locators\.([a-zA-Z0-9_]+)\)', code)

# Get unique locator names
unique_locators = sorted(set(pattern))

# Print results
print("Locators used in the file:")
for loc in unique_locators:
    print(loc)
