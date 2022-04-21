def print_data(user):
    match user:
        case ("Tom", 37):
            print("default user")
        case ("Tom", age):
            print(f"Age: {age}")
        case (name, 22):
            print(f"Name: {name}")
        case (name, age):
            print(f"Name: {name}  Age: {age}")
 
 
print_data(("Tom", 37))     # default user
print_data(("Tom", 28))     # Age: 28
print_data(("Sam", 22))     # Name: Sam
print_data(("Bob", 41))     # Name: Bob  Age: 41