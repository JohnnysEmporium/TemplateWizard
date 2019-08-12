import intestines

while True:
    a = input('Select Number\n1. Create New\n2. Change Color of existing one\n3. Exit\n---\n')
    if a == '1':
        intestines.paster()
    elif a == '2':
        color = input('Select Color\n1. Red\n2. Orange\n3. Green\n4. Add latest update to Notification\n---\n')
        print(color)
        if color != '1' and color != '2' and color != '3' and color != '4':
            print('\n!!!---You must input a number between 1 and 3---!!!\n')
        else:
            intestines.colors(color)
    elif a == '3':
        break
    else:
        print('\n!!!---You must input a number between 1 and 3\---!!!n') 
    
    