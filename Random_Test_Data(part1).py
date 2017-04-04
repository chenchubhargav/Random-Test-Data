import numpy as np
import pandas as pd
import random
import string

writer = pd.ExcelWriter('Output(part1).xlsx')

states = ['MO', 'IN', 'TN', 'TX', 'NY', 'OH', 'AL', 'AK', 'AZ', 'CT']
age = [1, 2, 3, 4, 5, 6, 7, 8, 9]
firstname = ['Kiran', 'Pavan', 'Keerthi', 'Sumanth', 'Hima Bindu', 'Vivek', 'Pranay', 'Murali', 'Disha', 'Bhavya', 'Lakshmi', 'Swaroop', 'Lokesh', 'Sridevi', 'Anuja', 'Vinay', 'Mark', 'James', 'Shravya', 'Vamsi']
lastname= ['Thota', 'Goti', 'Indela', 'Valligetla', 'Babu', 'Kesireddy', 'Agarwal', 'Meruva', 'Kohli', 'Teja', 'Chenchu', 'Valligetla', 'Katari', 'Mabbu', 'Bhave', 'Kothapally', 'Antony', 'Tally', 'Peddi', 'Mendu']

d = {'Kiran': 'male', 'Pavan': 'male', 'Sumanth': 'male', 'Vivek': 'male', 'Pranay': 'male', 'Murali': 'male', 'Bhavya': 'male', 'Swaroop': 'male', 'Lokesh': 'male', 'Vinay': 'male', 'Mark': 'male', 'James': 'male', 'Vamsi': 'male',
     'Keerthi': 'female', 'Hima Bindu': 'female', 'Disha': 'female', 'Lakshmi': 'female', 'Sridevi': 'female', 'Anuja': 'female', 'Shravya': 'female'}

G = []
rd = pd.DataFrame({})
rd['First Name'] = np.random.choice(firstname, 10000, replace=True)
fn = rd['First Name'].tolist()
for i in range(len(fn)):
    G1 = d[fn[i]]
    G.append(G1)
rd['Last Name'] = np.random.choice(lastname, 10000, replace=True)
rd['Gender'] = G
rd['Region'] = np.random.choice(states, 10000, replace=True)
rd['Age'] = np.random.choice(age, 10000, replace=True)
CUSID = list()
LN = list()
PR = list()
OW = list()
IN = list()
for i in range(len(rd)):
    c2 = str(random.choice(string.ascii_uppercase)) + str(random.randint(0, 9)) + str(random.randint(0, 9)) + str(
        random.randint(0, 9)) + str('-') + str(random.choice(string.ascii_lowercase)) + str(random.randint(0, 9))
    CUSID.append(c2)
    L = float('{0:.2f}'.format(random.uniform(1, 9999999)))
    LN.append(L)
    P = float('{0:.2f}'.format(random.uniform(1, L)))
    PR.append(P)
    O = L - P
    OW.append(O)
    I = O*(20/100)
    IN.append(I)

rd['CustomerID'] = CUSID
rd['Loan'] = LN
rd['Principal'] = PR
rd['Owed'] = OW
rd['Interest Due'] = IN
print(rd)

rd.to_excel(writer,'Sheet1')
writer.save()
writer.close()

#print(rd)
