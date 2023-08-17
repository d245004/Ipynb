import pandas as pd

i_file = "C:/Users/Jaeri/Downloads/3년 무이동 재고 (20230817).xls"
imsi = pd.read_excel(i_file,usecols='b,c,d,e,g,h,i,q')
str_expr = "M_LOC.str.startswith('12W06')" # Tiger로 시작하는 문자열

df_q = imsi.query(str_expr)    
df_q.to_excel("C://Users//Jaeri//Downloads//3년 무이동 12W-06 재고.xlsx",index=True)


# df_q.to_excel("C://Users//Jaeri//Downloads//3년 무이동 12W 재고.xlsx",index=False)

# imsi['location'] = imsi['M LOC'].str[:3]

# imsi.rename(columns={'M LOC':'M_LOC'},inplace=True)  #열 지정하여 이름 변경
# imsi = imsi.fillna('')

# imsi = imsi[imsi['M_LOC'] == '']
# imsi['M_LOC'] = imsi['M LOC']

# imsi.groupby("location")["L"].count()

# imsi[imsi['location'] == '12W01']
# ABC = imsi.query('M_LOC in "12W02"' )
# df_q
# imsi
# 컬럼명.str.startswith(비교문자열)
# imsi.query('M_LOC.str.startswith("09T") and L in "H" and CLASS.str.startswith("C")')

# imsi.query('M_LOC in "" and L in "H"')
# imsi.query(' location in "" ')
