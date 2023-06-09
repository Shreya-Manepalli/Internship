import re


t1 = 'Evaluate the following equation y = mx + c, where m is slope and c is intercept'
t2 = 'Rearrange the following: overflow stack'
t3 = 'Find x if 2^3 = 2x'
t4 = 'Find x if {(2x+3)* (23x+3)} = 256'


math_patterns = (r'((y|x)\s*=\s*.+),', r'if\s(.+)', r'({.+)')


for pattern in math_patterns:

    if re.search(pattern, t1):
        print(re.search(pattern, t1).group(1))
        continue

    if re.search(pattern, t3):
        print(re.search(pattern, t3).group(1))
        continue

    if re.search(pattern, t4):
        print(re.search(pattern, t4).group(1))
        continue