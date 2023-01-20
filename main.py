# SW 프로그래밍 기말 프로젝트 - 단순선형회귀분석 시스템
# 지구시스템과학과 2020134020 장보경
"""
****************** 분석 전 엑셀 파일은 꼭 현재 작업 위치에 저장하기! (main.py 있는 위치) ******************
<분석할 파일 1>
파일 이름: 월별 미세먼지농도와 국내여행횟수 (2018.1~2020.12).xlsx
설명변수 이름: 미세먼지농도 (PM 2.5)
반응변수 이름: 국내여행횟수
산점도의 제목: 월별 미세먼지농도와 국내여행횟수
"""
"""
<분석할 파일 2>
파일 이름: 생활물가지수와 출산율(2010~2020).xlsx
설명변수 이름: 총지수
반응변수 이름: 합계출산율
산점도의 제목: 생활물가지수와 합계출산율
"""

import os
import openpyxl
import pandas as pd
import numpy as np
import matplotlib
import matplotlib.pyplot as plt
import scipy.stats as stats
from sklearn.linear_model import LinearRegression


class CallData:
    def __init__(self, name):
        self.name = name

    def readExcel(self):            # 엑셀 읽어오기
        df = pd.read_excel(self.name)
        print(df)


class ScatterPlot(CallData):
    def __init__(self, name, x_axis, y_axis, title):
        super().__init__(name)
        self.x_axis = x_axis
        self.y_axis = y_axis
        self.title = title

    def getX(self, x_axis):         # 설명변수 데이터만 가져오기
        df = pd.read_excel(self.name)
        colX = df[x_axis]
        return colX

    def getY(self, y_axis):         # 반응변수 데이터만 가져오기
        df = pd.read_excel(self.name)
        colY = df[y_axis]
        return colY

    def plot(self):         # 산점도 그리기
        df = pd.read_excel(self.name)
        matplotlib.rcParams['font.family'] = 'Malgun Gothic'
        matplotlib.rcParams['axes.unicode_minus'] = False

        plt.scatter(df[x], df[y], c = 'green')
        plt.title(self.title)
        plt.xlabel(x)           # x축 이름 지정
        plt.ylabel(y)           # y축 이름 지정
        plt.show()

    def plotWithLine(self, a, b):           # 산점도 + 회귀식 그리기
        df = pd.read_excel(self.name)
        matplotlib.rcParams['font.family'] = 'Malgun Gothic'
        matplotlib.rcParams['axes.unicode_minus'] = False           # 산점도 상의 한글 꺠짐 방지

        plt.scatter(df[x], df[y], c='green')
        plt.title(self.title)
        plt.xlabel(x)  # x축 이름 지정
        plt.ylabel(y)  # y축 이름 지정

        plt.plot(df[x], a * df[x] + b, color = 'red')
        plt.show()


class SimpleLinear(ScatterPlot):
    def __init__(self, name, x_axis, y_axis, title):
        super().__init__(name, x_axis, y_axis, title)
        self.x_list = self.getX(x_axis)
        self.y_list = self.getY(y_axis)

    def test(self, a, b):
        predict_y = []
        for i in range(len(self.x_list)):
            predict_y.append(self.x_list[i] * a + b)         # 추정 회귀식에 x 변수를 대입했을 때의 y값을 배열에 저장하기
        y_mean = np.mean(self.y_list)            # y 변수의 평균

        # MSR 구하기
        msr = 0
        for i in range(len(predict_y)):
            msr += (predict_y[i] - y_mean) * (predict_y[i] - y_mean)

        # MSE 구하기
        mse = 0
        for i in range(len(predict_y)):
            mse += (self.y_list[i] - predict_y[i]) * (self.y_list[i] - predict_y[i])
        mse = mse / (len(self.y_list) - 2)

        # F통계량 구하기
        f = msr / mse
        alpha = 0.01            # 유의수준 0.01
        critical_value = stats.f.ppf(1-alpha, 1, len(self.y_list)-2)
        linearity = False
        if f >= critical_value:
            print(f'F통계량은 {f}로, 두 변수 간에는 선형적인 관계가 있습니다.\n')
            linearity = True
        else:
            print(f'F통계량은 {f}로, 두 변수 간에는 선형적인 관계가 없습니다.\n')

        # 결정계수 구하기 (F검정에서 선형성이 인정되었을 경우에만)
        if linearity:
            ssr = msr
            sse = mse * (len(self.y_list) - 2)
            sst = ssr + sse
            r_2 = ssr / sst
            print(f'결정계수는 {r_2} 입니다.\n')


# Step 1. 분석할 파일 이름.xlsx 입력받기

print('Step 1. 분석할 파일을 가져옵시다.\n')
file = input('분석할 파일이름.xlsx를 입력하세요: ')          # 불러오기 전 현재 작업 위치에 엑셀 파일 저장하기
print('\n')
data = CallData(file)
data.readExcel()
print('-------------------------------------------------------------------------------------------------\n')

# Step 2. 산점도 그리기

print('Step 2. 산점도를 그려봅시다.\n')
x = input('설명변수가 될 x축 데이터의 이름을 입력하세요: ')
y = input('반응변수가 될 y축 데이터의 이름을 입력하세요: ')
my_title = input('산점도의 제목을 입력하세요: ')
sp = ScatterPlot(file, x, y, my_title)
sp.plot()
print('-------------------------------------------------------------------------------------------------\n')

# Step 3. 단순선형회귀식 추정하기 (y = b_0 + b_1 * x)

print('Step 3. 단순선형회귀식을 추정합니다.\n')
myData = SimpleLinear(file, x, y, my_title)

line = LinearRegression()
line.fit(myData.x_list.values.reshape(-1, 1), myData.y_list)          # x가 2차원 array이기 때문에 my_x.values.reshape(-1, 1)
b_1 = line.coef_[0]
b_0 = line.intercept_
print(f'추정된 단순선형회귀식은 y = {b_1}x + {b_0} 입니다.\n')
myData.plotWithLine(b_1, b_0)           # 산점도에 추정 회귀식 같이 그리기

print('-------------------------------------------------------------------------------------------------\n')

# Step 4. 회귀모형의 적합성 판정

print('Step 4. 회귀모형이 유의한지 판단하기 위해 F통계량을 분석합니다.\n')

myData.test(b_1, b_0)
