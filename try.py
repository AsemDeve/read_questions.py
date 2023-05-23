import re

# target_string = """Ashleaf spots **
# Cutaneous neurofibromas
# Angiofibromas
# Shagreen patches
# Tinea corporis
# """
#
# # split on white-space
# word_list = re.split(r"\n+", target_string)
# print(word_list)

list1 = [1
    , 2
    , 3
    , 4
    , 5
    , 6
    , 7
    , 8
    , 9
    , 10
    , 11
    , 12
    , 13

         ]
list2 = ["Meanwhile",
         "Then",
         "After that",
         "Later",
         "Soon",
         "After awhile",
         "Next",
         "Second",
         "Third",
         "Secondly",
         "Thirdly",
         "And",
         "Furthermore",
         "Further",
         "Moreover",
         "Another"
         ]
#print(list1)

#print(list2)
for question, choices in zip(list2, list1):
    print(question)
    print(choices)