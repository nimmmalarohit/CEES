from base64 import b64encode
from time import sleep

actual_password = input("Enter your password: ")
print("Below is your encoded password: \n")
print(b64encode(actual_password.encode()))
with open('encoded_password_file.txt', 'w') as pass_file:
    pass_file.write(str(b64encode(actual_password.encode())))