import pandas as pd

rd = pd.read_excel('Output(part1).xlsx')
#####################################################################################
#What is the mean loan size per region?
df1 = pd.DataFrame({})
By_Region = rd.groupby('Region').mean()
By_Region_Loan = By_Region.loc[:,'Loan']
df1['By Region Loan'] = By_Region_Loan
#print(df1)
#print(By_Region_Loan)
####################################################################################
#What is the mean loan size per gender?
df2 = pd.DataFrame({})
By_Gender = rd.groupby('Gender').mean()
By_Gender_Loan = By_Gender.loc[:,'Loan']
df2['By Gender Load'] = By_Gender_Loan
#print(df2)
#print(By_Gender_Loan)
####################################################################################
#What is the mean loan size per age group?
df3 = pd.DataFrame({})
By_Age = rd.groupby('Age').mean()
By_Age_Loan = By_Age.loc[:,'Loan']
df3['By Age Loan'] = By_Age_Loan
#print(df3)
#print(By_Age_Loan)
###################################################################################
#What is the maximum amount owed by region for each age group?
df4 = pd.DataFrame({})
Owed_By_Region_Age = rd.groupby(['Region', 'Age'])['Owed'].max()
df4['Owed By Region Age'] = Owed_By_Region_Age
#print(df4)
#print(Owed_By_Region_Age)
##################################################################################
#What is the maximum amount owed by gender for each age group?
df5 = pd.DataFrame({})
Owed_By_Gender_Age = rd.groupby(['Gender', 'Age'])['Owed'].max()
df5['Owed By Gender Age'] = Owed_By_Gender_Age
#print(Owed_By_Gender_Age)
#################################################################################
#How much interest can the company expect to make from each region?
df6 = pd.DataFrame({})
Interest_By_Region = rd.groupby('Region')['Interest Due'].sum()
df6['Interest By Region'] = Interest_By_Region
#print('df6')
#print(Interest_By_Region)
#################################################################################
#Which man in each region owes the most?
L =[['Region', 'Name']]
Amtowes_By_Region = rd[rd['Gender'] == 'male'].groupby('Region')['Owed'].max()
for i in Amtowes_By_Region:
    for j,k,l,m in zip(rd['Owed'], rd['First Name'], rd['Last Name'], rd['Region']):
        if i == j:
            ABR = m+'  '+k+' '+l
            ABR = ABR.split('  ')
            L.append(ABR)
            #print(ABR)
#print(L)
df7 = pd.DataFrame(L[1:], columns=L[0])
#print(df7)
################################################################################


writer = pd.ExcelWriter('Output(part2).xlsx')
df1.to_excel(writer,'Sheet1')
df2.to_excel(writer,'Sheet2')
df3.to_excel(writer,'Sheet3')
df4.to_excel(writer,'Sheet4')
df5.to_excel(writer,'Sheet5')
df6.to_excel(writer,'Sheet6')
df7.to_excel(writer,'Sheet7')

writer.save()
writer.close()