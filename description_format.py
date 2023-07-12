'''
import re

def extract_values(statement):
    match = re.search(r"((\d+\.\d+)\sor\s(\d+\.\d+))", statement)

    if match:
        measured_value = float(match.group(2))
        deviation_value = float(match.group(3))
        return  measured_value, deviation_value
    else:
        return None

# Example usage
statement = "True Position |0.1|A|B|C| is  0.154 or 0.054 out of tolerance Please see attached 44DAC02057 v2  support pack  for details."
values = extract_values(statement)
if values:
    measured_value, deviation_value = values
    print("Measured Value:", measured_value)
    print("Deviation Value:", deviation_value)
else:
    print("No match found.")

'''
import re

def extract_values(statement):
    match1 = re.findall(r'is(.*?)or', statement)
    match2 = re.findall(r'or(.*?)', statement)
    return match1,match2
# Example usage
statement = "True Position |0.1|A|B|C| is  0.154 or 0.054 out of tolerance Please see attached 44DAC02057 v2  support pack  for details."
values = extract_values(statement)
if values:
    measured_value = values
    print("Measured Value:", measured_value)
else:
    print("No match found.")
