import re

sentences = [
    "Diameter 645. +/- 0.1mm is 644.86  or  0.04 under minimum.Please see attached 44DAC02057 v2  support pack  fordetails.",
    "True Position |0.1|A|B|C| is  0.154 or 0.054 out of tolerance Please see attached 44DAC02057 v2  support pack  for details.",
    "4 pos True Position 0.25|A|B|C| is  0.628 or 0.378 out of tolerance Please see attached 44DAC02057 v2  support pack  for details."
]

# Regular expression pattern to extract the measured value and deviation value
pattern = r"is\s+(.*?)\s+or\s+(\d+\.?\d*)"

measured_values = []
deviation_values = []

for sentence in sentences:
    if sentence.startswith("Casting Blend"):
        measured_value = ""
        deviation_value = ""
    else:
        # Search for the pattern in the sentence
        match = re.search(pattern, sentence)

        if match:
            measured_value = match.group(1)
            deviation_value = match.group(2)
        else:
            measured_value = ""
            deviation_value = ""

    measured_values.append(measured_value)
    deviation_values.append(deviation_value)

# Print the measured values and deviation values with a space in between
print("Measured values:", " ".join(measured_values))
print("Deviation values:", " ".join(deviation_values))


