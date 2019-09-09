import intestines

while True:
    
    a = input('''
Select Number
---------------
1. Create New                                    --> Saves to chosen file
2. Change Color of existing one or add update    --> Reads and saves to chosen file 
3. Exit
---------------
''')
    
    
    if a == '1':
        intestines.paster()
        
    elif a == '2':
        
        color = input('''
Select Option
---------------
1. Change Color
2. Add Update
---------------
''')
        
        if color != '1' and color != '2':
            print('\n!!!---You must input a number between 1 and 2---!!!\n')
        
        elif color == '1':
            color2 = color + input('''
Select Colour
---------------
1. Red
2. Orange
3. Green
---------------
''')
            if color2 != '11' and color2 != '12' and color2 != '13':
                print('\n!!!---You must input a number between 1 and 3---!!!\n')
            else:
                intestines.colors(color2)
                
        elif color == '2':
            color3 = color + input('''
Select Option
---------------
1. Update Red
2. Update Orange
3. Update Green
4. Only Update
---------------
''') 
            
            if color3 != '21' and color3 != '22' and color3 != '23' and color3 != '24':
                print('\n!!!---You must input a number between 1 and 3---!!!\n')
            else:
                intestines.colors(color3)
            
        else:
            intestines.colors(color)
    
    elif a == '3':
        break
    
    else:
        print('\n!!!---You must input a number between 1 and 3---!!!\n') 
    
    