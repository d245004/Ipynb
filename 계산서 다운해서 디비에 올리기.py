import pandas as pd
source_file = "C:/Users/Jaeri/Downloads/maip.xls"
work_file = pd.read_excel(source_file,usecols= 'a,b,e,g,h,j,l,m,o,p,q,u',header=6)
work_file.to_csv("C:/Users/Jaeri/Downloads/maip.csv",index=False)

source_file_1 = "C:/Users/Jaeri/Downloads/machul.xls"
work_file_1 = pd.read_excel(source_file_1,usecols= 'a,b,e,g,h,j,l,m,o,p,q,u',header=6)
work_file_1.to_csv("C:/Users/Jaeri/Downloads/machul.csv",index=False)




# work_file