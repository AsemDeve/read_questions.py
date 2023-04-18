import re

target_string = """Ashleaf spots **
Cutaneous neurofibromas 
Angiofibromas
Shagreen patches 
Tinea corporis
"""

# split on white-space
word_list = re.split(r"\n+", target_string)
print(word_list)
