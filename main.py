import itertools
import win32com.client as client
import time
from string import digits, punctuation, ascii_letters
from datetime import datetime


symbols = digits + punctuation + ascii_letters


def brute_exel_doc():
    print("Hello this is program to crack a password of excel document")
    
    try:
        password_length = input("Enter password length, from time to time, for example 3 - 7: ")
        password_length = [int(item) for item in password_length.split("-")]
    except:
        print("Please check the input data")
        
    print("If password contains only digits, enter: 1\nIf password contains only letters, enter: 2\n"
          "If passowrd contains letters and digits, enter: 3\nIf password contains also special symbols, enter: 4")
    
    try:
        choise = int(input(""))
        if choise == 1:
            possible_symbols = digits
        elif choise == 2:
            possible_symbols = ascii_letters
        elif choise == 3:
            possible_symbols = digits + ascii_letters
        elif choise == 4:
            possible_symbols = digits + ascii_letters + punctuation
        else:
            print("Number of out range")
        
    except:
        print("Please check the input data")
    
    
    start_timestamp = time.time()
    print(f"Started at {datetime.utcfromtimestamp(time.time()).strftime('%H:%M:%S')}")
    count = 0
    for pass_length in range(password_length[0], password_length[1]+1):
        for password in itertools.product(possible_symbols, repeat=pass_length):
            password = "".join(password)
            
            opened_doc = client.Dispatch("Excel.Application")
            count += 1
            try:
                opened_doc.Workbooks.Open(
                    r"D:\Programming\Python\Freelance\crack_exel\test.xlsx", 
                    False, 
                    True,
                    None,
                    password
                    )
                print(f"Finished at {datetime.utcfromtimestamp(time.time()).strftime('%H:%M:%S')}")
                print(f"Time {time.time() - start_timestamp}")
                print(f"Attempt {count} Password {password}")
                
                return 
            except:
                if count % 10000 ==0:
                    print(f"Atempt {count} Password {password}")
                


if __name__ == '__main__':
    brute_exel_doc()