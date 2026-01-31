c=0
k=201104

bimestreactual=4

y=[]
for i in range(1,86):

    k+=2    
    print(k)


    if (k-201100-100*c)==14:
        c+=1 
        k+=100
        k-=12


    y.append(k)

print(y)
print(len(y))