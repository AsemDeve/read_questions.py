import pandas as pd


check_image_list = ['image', ' picture', 'shown', ' figure', ' graph', ' diagram', ' table', 'radiograph', 'DVH',
                    ' CT ', 'x-ray',' OCT']
list_of_questions_with_images = {}

df = pd.read_excel('doc.xlsx')

# Iterate over the second column and print the values
for index, value in df.iloc[:, 1].items():
    for word in check_image_list:
        if word.lower() in value.lower():