def factorial(a):
    if(a==1):
        return 1
    return factorial(a-1)*a
a=int(input("Enter a number for finding the factorial..."))
fact=factorial(a)

print(fact)
