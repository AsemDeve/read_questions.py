import docx

# Replace 'file_name.docx' with the name of your file
doc = docx.Document('Radio3.docx')

for para in doc.paragraphs:
    # Split the paragraph text into question and answer choices
    question, choices = para.text.split('\n', 1)

    # Print the question and answer choices
    print('Question:', question)
    print('Choices:', choices)

    # Split the answer choices into individual options
    options = choices.split('\n')[1:]

    # Print each answer option with its corresponding letter
    for i, option in enumerate(options):
        print(chr(ord('a') + i) + ')', option.strip())

    # Add a line break between each question
    print()
