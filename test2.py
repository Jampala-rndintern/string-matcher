import pandas as pd
import os
from fuzzywuzzy import fuzz
import nltk as dist



  
def comparitor(p360,SFL,p360_rct,sfl_rtc):
    k=0
    i=0
    j=0
    count=0
    check=[]
    presense=[]
    per_count=[]
    per=[]
    SFL_sto=[]
    p360_bckt=[]
    sfl_bckt=[]
    for i in range(len(p360)):
        p360[i]=("".join(filter(str.isalnum, p360[i]))).casefold()
        p360_bckt.append(p360[i]+"            "+str(i))
    p360.sort()
    for i in range(len(SFL)):
        SFL[i]=("".join(filter(str.isalnum, SFL[i]))).casefold()
        sfl_bckt.append(SFL[i]+"            "+str(i))
    SFL.sort()
    sfl_bckt.sort()
    i=0
    while j < len(SFL):
        if i==len(p360):
            break
        string=p360[i]
        ser_str=SFL[j]
        #print(string)
        #print(ser_str,"\n")
        if k==0 and string[k]==ser_str[k] :                 #and k <min( len(string),len(ser_str))
            k+=1
            j-=1
        elif k < min( len(string),len(ser_str)) and string[k]==ser_str[k] and string[:k-1] == ser_str[:k-1]  :
            k+=1
            j-=1
        elif k < min( len(string),len(ser_str)) and string[k]<ser_str[k] and string[:k-1] == ser_str[:k-1] :
            k=0
            j-=1
            # print(p360[i], "                            object not found")
            presense.append("no")
            per_count.append(dist.edit_distance(ser_str,string))
            per.append(fuzz.ratio(string, ser_str))
            size_temp=len(ser_str)+1
            num=int(((sfl_bckt[j+1])[size_temp:]))
            SFL_sto.append(sfl_rtc[num])
            j-=1
            if per_count[i]<=4:
                check.append("!!! probably found !!!")
                count+=1
            else:
                check.append("likely not found")
            i+=1
        elif k < min( len(string),len(ser_str)) and  string[:k-1] != ser_str[:k-1]  and k>0:        #string[k]>=ser_str[k] and
            k=0
            #print(p360[i], "                            object not found")
            presense.append("no")
            per_count.append(dist.edit_distance(ser_str,string))
            per.append(fuzz.ratio(string, ser_str))
            #SFL_sto.append(ser_str)
            size_temp=len(ser_str)+1
            num=int(((sfl_bckt[j])[size_temp:]))
            SFL_sto.append(sfl_rtc[num])
            if per_count[i]<=4 or per[i]>75:
                check.append("!!! probably found !!!")
                count+=1
            else:
                check.append("likely not found")
            i+=1
            j-=1
        else:
            k=0    
        if k== min( len(string),len(ser_str)) and string[:k-1] == ser_str[:k-1]:
           #print(p360[i],"                         object found")
            presense.append("yes")
            per_count.append(dist.edit_distance(ser_str,string))
            per.append(fuzz.ratio(string, ser_str))
            #SFL_sto.append(ser_str)
            size_temp=len(ser_str)+1
            num=int(((sfl_bckt[j+1])[size_temp:]))
            SFL_sto.append(sfl_rtc[num])
            if  per_count[i]==0:
                check.append("✔️perfect match")
                
            else:
                check.append("!!!matched but partially with slesforce")
            count+=1
            k=0
            i+=1
            
        #if k==len(string) and 
        j+=1

    
    p360_bckt.sort()
    
    for i in range (len(p360_bckt)):#
        size_temp=len(p360[i])+1
        num=int(((p360_bckt[i])[size_temp:]))
        p360[i]=p360_rct[num]
    #for i in range (len(SFL_sto)):#
       # size_temp=len(sfl[i])+1
      #  num=int(((sfl_bckt[i])[size_temp:]))
     #   SFL_sto[i]=sfl_rct[num]
    #print(p360)
    print("\n\n\n------------------total number of matches likely to be found",count,"--------------------------------- \n\n\n")

    df = pd.DataFrame({'Company Name': p360,'matching cmpny':SFL_sto, 'Presense': presense,'lvs dist':per_count,'percentage':per,'conclusion':check}) #
    
    #df=df.sort_values("Presense",ascending=False)
    
    writer = pd.ExcelWriter('demo.xlsx', engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Sheet1', index=False)        
    writer.save()
    os.system("start EXCEL.EXE demo.xlsx")
    
        

if __name__ == "__main__":
    df=pd.read_excel('./p360 client.xlsx',header = 0, converters={"Client Name":str}) 	# by changing it to request from the user like typping name and changings list terms to char and removing all the lst iterators we can generate suggestions
    #df.sort_values(by='Client Name')
    p360=df['Client Name'].tolist()
    p360_rct=df["Client Name"].tolist()
    df2=pd.read_excel("./salesforce client.xlsx", header=0,converters={"Account Name":str})
    #df2.sort_values(by='Account Name')
    SFL=df2['Account Name'].tolist()
    sfl_rtc=df2['Account Name'].tolist()
    #print(SFL)
    
    print("\n\n\n\n\n\n\nintgiation reading dne\n\n\n\n\n\n\n")
    comparitor(p360,SFL,p360_rct,sfl_rtc)
 
