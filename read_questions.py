import os
import re
import webbrowser
import docx2txt
import xlsx_operations
import pandas as pd

""" this code read questions and choices from microsoft word (.docx) file , split choices by new line """
question_pattern = r"(#.*?#)"
choices_pattern = '\n+'  # split choices by new line
choices_pattern_1 = "^[a-fA-F][.|\-|)|:]"
choices_pattern_2 = "^[(][a-fA-F][.|\-|)|:]"
choices_pattern_3 = "^[a-fA-F][ ]"
choices_pattern_4 = "^[\d][.|\-|)|:]"
choices_pattern_by_letter = choices_pattern_1 + '|' + choices_pattern_2 + '|' + choices_pattern_3 + '|' + choices_pattern_4
header = ['questionType', 'questionText', 'option1', 'option2', 'option3', 'option4', 'option5', 'answer']
check_image_list = ['image', ' picture', 'shown', ' figure', ' graph', ' diagram', ' table', 'radiograph', 'DVH',
                    ' CT ', 'x-ray', ' OCT']
directory = 'doc/'
for docxfile in os.scandir(directory):  ## iterate over files in directory
    if docxfile.is_file() and docxfile.name.endswith('.docx'):
        print('proccessing : ' + docxfile.name)
        doc = docx2txt.process(docxfile.path)
        matches = re.split(question_pattern, doc, flags=re.DOTALL)  # to split answers and questions :

        list_of_questions = matches[1::2]  # Elements from "matches" list starting from 1 iterating by 2
        list_of_choices = matches[2::2]  # Elements from "matches" list starting from 2 iterating by 2

        questions_with_images = {}
        question_and_choices = []
        question_and_choices.append(header)
        duplicated_questions = 0

        for question, choices in zip(list_of_questions, list_of_choices):
            if '$' in question:  # to ignore duplicated questions, while dublicate questions have ($) in it
                duplicated_questions += 1
            else:
                inner_list = ["mc"]
                removeH = re.sub(r"#", "", question)  # remove # from question Text
                cleanQ = re.sub(r"^(\d{1,3}[.|\-|)|:])", "", removeH.strip())  # remove question number
                inner_list.append(cleanQ.strip())
                raw_split_choices = re.split(choices_pattern, choices)
                split_choices = []
                for choice in raw_split_choices:  # remove empty choices from choices list
                    if choice.strip() != '':
                        split_choices.append(choice.strip())
                asterisk_count = 0
                for string in split_choices:  # check if no one of choices has answer or , more than one choice has answer
                    if "*" in string:
                        asterisk_count += 1
                        if asterisk_count > 1:
                            break
                if asterisk_count == 0 or asterisk_count > 1:
                    print(
                        "no answer found!!!! or More than one answer found!!!! in  question : " + "\n" + str(question))
                    exit()

                answer_index = 0
                for item in split_choices:
                    cleanChoice = re.sub(choices_pattern_by_letter, "", item.strip())  # remove choices numbering
                    if '*' in item:
                        answer_index = split_choices.index(item) + 1
                        cleanChoice = re.sub("\*", "", item)  # remove * from the answer.
                        cleanChoice = re.sub(choices_pattern_by_letter, "",
                                             cleanChoice.strip())  # remove choices numbering from the answer
                    if not cleanChoice.endswith(".") and cleanChoice.strip() != "":  # add (.) to the end of choices.
                        inner_list.append(cleanChoice.strip() + '.')
                    else:
                        inner_list.append(cleanChoice.strip())

                if len(inner_list) != len(set(inner_list)):  # check duplicates  in choices
                    print("choices duplicates found in Question: " + "\n" + str(question))
                    exit()
                inner_list.append(None)  # to keep list and answer index wrapped.
                inner_list.append(None)  # to keep list and answer index wrapped.
                inner_list.append(None)  # to keep list and answer index wrapped.
                inner_list.insert(7, answer_index)
                question_and_choices.append(inner_list)

        print("total number of questions: " + str(len(list_of_questions)) + " //marked duplicated questions : " + str(
            duplicated_questions))
        xlsx_operations.write('doc/' + docxfile.name[:-4] + 'xlsx', question_and_choices)

        # count number of choices for each question
        df = pd.read_excel('doc/' + docxfile.name[:-4] + 'xlsx')
        row_counts = df.count(axis=1)
        grouped_counts = row_counts.value_counts().reset_index()
        grouped_counts.columns = ['Number of Choices', 'Counts']
        grouped_counts['Number of Choices'] = grouped_counts['Number of Choices'] - 3
        print(grouped_counts)

        # details for images found in each question
        for index, value in df.iloc[:, 1].items():
            img_words_found = []
            for word in check_image_list:
                if word.lower() in value.lower():
                    img_words_found.append(word)
            if img_words_found:
                questions_with_images[index + 2] = img_words_found

        df_result = pd.DataFrame(questions_with_images.items(), columns=['Excel row', 'Words'])
        print('---------------------------------------')
        print('details for images found in each question')
        print(df_result)
        print('---------------------------------------')
        print('details for count of choices in each question')
        # details for count of choices in each question
        for index, row in df.iterrows():
            non_empty_choice_count = row.count()
            print(f"excel row  {index + 2} - number of choices: {non_empty_choice_count - 3}")

        cwd = os.getcwd()
        file_path = os.path.join(cwd, 'doc/' + docxfile.name[:-4] + 'xlsx')
        webbrowser.open('file://' + file_path)
