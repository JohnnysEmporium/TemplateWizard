import intestines

while True:
    a = input('''
    Select Number
    1. Create New                                    --> working on template.docx
    2. Change Color of existing one or add update    --> working on color_change.docx
    3. Exit
    ---
    ''')
    if a == '1':
        intestines.paster()
    elif a == '2':
        color = input('''
        Select Color
        1. Red
        2. Orange
        3. Green
        4. Add latest update to Notification
        ---
        ''')
        print(color)
        if color != '1' and color != '2' and color != '3' and color != '4':
            print('\n!!!---You must input a number between 1 and 3---!!!\n')
        else:
            intestines.colors(color)
    elif a == '3':
        break
    else:
        print('\n!!!---You must input a number between 1 and 3\---!!!n') 
    
    