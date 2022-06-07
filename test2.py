import pandas as pd

def comparitor(p360,SFL):
    k=0
    i=0
    j=0
    presense=[]
    while j < len(SFL):
        if i==len(p360):
            break
        string=p360[i]
        ser_str=SFL[j]
        while len(ser_str) < len (string):
            j+=1
            ser_str=SFL[j]

        #print(string)
        #print(ser_str,"\n")
  
        if k==0 and string[k]==ser_str[k] and k < len(string):
            k+=1
            j-=1

        elif string[k]==ser_str[k] and string[:k-1] == ser_str[:k-1] and k < len(string):
            k+=1
            j-=1
        elif string[k]<ser_str[k] and string[:k-1] == ser_str[:k-1] and k < len(string):
            k=0
            j-=1
            print(p360[i], "                            object not found")
            presense.append("no")
            i+=1    
        else:
            k=0    
           
        if k==len(string) and string[:k-1] == ser_str[:k-1]:
               print(p360[i],"                         object found")
               presense.append("yes")
               k=0
               i+=1
        #if k==len(string) and 
        j+=1
    writer = pd.ExcelWriter('demo.xlsx', engine='xlsxwriter')
    writer.save()
    df = pd.DataFrame({'Company Name': p360,
                   'Presense': presense})
    writer = pd.ExcelWriter('demo.xlsx', engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    writer.save()
    
        

if __name__ == "__main__":
    df=pd.read_excel('./p360 client.xlsx',header = 0, converters={"Client Name":str}) 	# by changing it to request from the user like typping name and changings list terms to char and removing all the lst iterators we can generate suggestions
    #df.sort_values(by='Client Name')
    p360=df['Client Name'].tolist()
    p360.sort()
    df2=pd.read_excel("./salesforce client.xlsx", header=0,converters={"Account Name":str})
    #df2.sort_values(by='Account Name')
    SFL=df2['Account Name'].tolist()
    SFL.sort()
    #print(SFL)
    comparitor(p360,SFL)
 
