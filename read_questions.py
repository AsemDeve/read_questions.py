import os
import re
import webbrowser
import write_csv
import docx2txt

doc = docx2txt.process('test.docx')

question_pattern = r"(#.*?#)"
choices_pattern = '[a-fA-F][.|-|)|:].*'

matches = re.split(question_pattern, doc, flags=re.DOTALL)  # to split answers and questions :

list_of_questions = matches[1::2]  # Elements from list1 starting from 1 iterating by 2
list_of_choices = matches[2::2]  # Elements from list1 starting from 2 iterating by 2
my_res = []

for question, choices in zip(list_of_questions, list_of_choices):
    inner_list = ["mc"]
    removeH = re.sub(r"#", "", question)  # remove # from question Text
    cleanQ = re.sub(r"^(\d{1,3}[-|.])", "", removeH.strip())  # remove question number
    inner_list.append(cleanQ)
    split_choices = re.findall(choices_pattern, choices)
    answer_index = 0
    for item in split_choices:
        cleanChoice = re.sub("^[a-fA-F][.|)|:]", "", item)
        inner_list.append(cleanChoice)
        if '*' in item:
            answer_index = split_choices.index(item) + 1
    inner_list.append("")
    inner_list.append("")
    inner_list.insert(7, answer_index)
    my_res.append(inner_list)

print("total number of question Extracted : " + str(len(list_of_questions)))

write_csv.writer('import_q.csv', my_res)
cwd = os.getcwd()
file_path = os.path.join(cwd, 'import_q.csv')
webbrowser.open('file://' + file_path)
