import intestines

while True:
    a = input('Select Number\n1. Create New\n2. Change Color of existing one\n3. Exit\n')
    if a == '1':
        intestines.paster()
        print('\nNotification created, check output.docx\n')
    elif a == '2':
        color = input('\nSelect Color\n1. Red\n2. Orange\n3. Green\n')
        print(color)
        if color != '1' and color != '2' and color != '3':
            print('\nYou must input a number between 1 and 3\n')
        else:
            intestines.colors(color)
            print('\nNotification created, check output.docx\n')
    elif a == '3':
        break
    else:
        print('\nYou must input a number between 1 and 3\n') 
    
    
    