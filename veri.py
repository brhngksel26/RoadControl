import numpy as np
def entropy(dist):
    su=0
    for p in dist:
        r= p/sum(dist)
        if r==0:
            su+=0
        else:
            su+= -r*(np.log(r))
    return su/np.log(2)

a = 1 / 3
b = 1 / 6

k = 3/5
l = 2/5

list = [0.5 , a , b]
list_b = [k , l]

x = entropy(list)
y = entropy(list_b)


print(y)