import intestines

while True:
    a = input('Select Number\n1. Create New\n2. Change Color of existing one\n3. Exit')
    if a == '1':
        intestines.paster()
        print('Notification created, check output.docx')
    elif a == '2':
        color = input('Select Color\n1. Red\n2. Orange\n3. Green\n')
        if color != '1' or color != '2' or color != '3':
            print('You must input a number between 1 and 3')
        else:
            intestines.colors(color)
            print('Notification created, check output.docx')
    elif a == '3':
        break
    else:
        print('You must input a number between 1 and 3') 
    
    
    