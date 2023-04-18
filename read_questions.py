import os
import re
import webbrowser
import docx2txt
import xlsx_operations

docx_name = 'ent'
doc = docx2txt.process('doc/' + docx_name + '.docx')
question_pattern = r"(#.*?#)"
choices_pattern = '\n+'  # split choices by new line
choices_pattern_1 = "^[a-fA-F][.|\-|)|:]"
choices_pattern_2 = "^[(][a-fA-F][.|\-|)|:]"
choices_pattern_by_letter=choices_pattern_1 + '|' + choices_pattern_2
header = ['questionType', 'questionText', 'option1', 'option2', 'option3', 'option4', 'option5', 'answer']

matches = re.split(question_pattern, doc, flags=re.DOTALL)  # to split answers and questions :

list_of_questions = matches[1::2]  # Elements from "matches" list starting from 1 iterating by 2
list_of_choices = matches[2::2]  # Elements from "matches" list starting from 2 iterating by 2

question_and_choices = []
question_and_choices.append(header)
duplicated_questions=0
for question, choices in zip(list_of_questions, list_of_choices):
    if '$' in question:  # to ignore duplicated questions
        duplicated_questions+=1
    else:

        inner_list = ["mc"]
        removeH = re.sub(r"#", "", question)  # remove # from question Text
        cleanQ = re.sub(r"^(\d{1,3}[.|\-|)|:])", "", removeH.strip())  # remove question number
        inner_list.append(cleanQ.strip())
        split_choices = re.split(choices_pattern, choices)
        answer_index = 0
        for item in split_choices:
            cleanChoice = re.sub(choices_pattern_by_letter, "", item.strip())  # remove choices numbering

            if not cleanChoice or cleanChoice == " ":
                pass  # this to ignore empty lines.
            else:
                if '*' in item:
                    answer_index = split_choices.index(item)
                    cleanChoice = re.sub("\*", "", item)  # remove * from the answer.
                    cleanChoice = re.sub(choices_pattern_by_letter, "",
                                         cleanChoice.strip())  # remove choices numbering from the answer

                if cleanChoice.endswith("."):
                    inner_list.append(cleanChoice.strip())
                else:
                    inner_list.append(cleanChoice.strip() + '.')  # add (.) to the end of choice if not exist.

        inner_list.append(None)  # to keep list and answer index wrapped.
        inner_list.append(None)  # to keep list and answer index wrapped.
        inner_list.append(None)  # to keep list and answer index wrapped.
        inner_list.insert(7, answer_index)
        question_and_choices.append(inner_list)

print("total number of questions : " + str(len(list_of_questions)))
print("number of duplicated questions : " +str(duplicated_questions))
xlsx_operations.write('doc/' + docx_name + '.xlsx', question_and_choices)

cwd = os.getcwd()
file_path = os.path.join(cwd, 'doc/' + docx_name + '.xlsx')
webbrowser.open('file://' + file_path)
