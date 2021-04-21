# -*- coding: utf-8 -*-
"""
Created on Mon Feb  1 10:05:35 2021

@author: User
"""

import os
path = 'C:/CodeR'
os.chdir(path)

#import functions
import pandas as pd
import datetime
import numpy as np

#%%
weights = pd.read_excel('KOSPI 월간 포트폴리오(SalesG 제외) 02.01.xlsx')
weights = weights.iloc[:,2:]                          # 아래, 오른쪽
weights = weights[~weights.iloc[:,2].isna()]          # 종목코드열 빈칸 제거
weights.columns = weights.iloc[0]                     # 4행을 columns로 저장
weights = weights[~(weights.iloc[:,1] == 'RANK')]     # columns로 저장해놔서 Rank 있는 행 제거
weights = weights.fillna(method='ffill')              # 앞에 날짜 있으면 앞에 날짜로 저장
weights.index = weights.iloc[:,0]                     # 날짜를 인덱스로 지정
weights.index.name = 'date'                           # 인덱스의 이름을 'date'라 명명
weights = weights.iloc[:,1:]                          # C열 버리고 D열부터 weights에 저장

# iloc -> dataframe을 index로 로케이션
# loc ->


#%%
rebal = sorted(list(set(list(weights.index))))        # list로 들고 와서 set으로 서로 다른 원소만 존재하게 됨(set은 순서가 없음)
codes = list(set(list(weights['종목코드'])))

# func = functions.funcs()
# price = func.getPrice(codes, str(dates[0])[:10])
price = pd.read_pickle('price.pkl')  # 기존에 쓰던 데이터 가져오기 (수정주가)
# price = func.updatePrice(codes, price)  # 업데이트 (자동 저장됨)

#%%
dates = list(price.index[price.index >= rebal[0]])    #
    
initial = 100000000000

result = []
ports = []
for i in range(len(dates)):      # range 1547까지
    date = dates[i]
    if date in rebal:   
        port = weights.loc[date][['종목코드', '진입주가', '종가', '비중']]                #loc (date로 index하여 가져올거 가져온다)
        port = pd.merge(port,price.loc[date],left_on='종목코드',right_index=True)      # left(port), right(index)
        port.columns = ['종목코드', '진입주가', '종가', '비중', '진입종가']                 # price를 진입종가로 붙임
       
        if len(result) == 0:
            port['주수'] = round(initial*port['비중']/port.iloc[:,-1]/100,0)           # 포트 구성, '진입종가'뒤에 '주수"붙는다
            temp = pd.merge(port,price.loc[date],left_on='종목코드',right_index=True)
            cash = initial - (temp.iloc[:,-1]*temp.iloc[:,-2]).sum()                  # '진입종가'와 '주수' 곱해서 캐시구한다

        else:
            ex = pd.merge(ports[-1],price.loc[date],left_on='종목코드',right_index=True)
            ex_nav =  (ex.iloc[:,-1]*ex.iloc[:,-2]).sum() + result[-1][-1]
            port['주수'] = round(ex_nav*port['비중']/port.iloc[:,-1]/100,0)
            temp = pd.merge(port,price.loc[date],left_on='종목코드',right_index=True)
            cash = ex_nav - (temp.iloc[:,-1]*temp.iloc[:,-2]).sum()
        
        ports.append(port)                                                            # 시계열로 계속 쌓이게

    else:
        cash = result[-1][-1]        
        
    temp = pd.merge(port,price.loc[date],left_on='종목코드',right_index=True)          # else에 넣어도 된다
    stocks = (temp.iloc[:,-1]*temp.iloc[:,-2]).sum()

    result.append([stocks+cash, stocks, cash])
    
result = pd.DataFrame(result)
result.index = dates

#%%

result.to_excel('result2.xlsx')


