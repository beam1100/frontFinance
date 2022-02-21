import pandas as pd
import matplotlib.pyplot as plt
plt.rcParams['axes.unicode_minus'] = False      # 맷플롯립 한글 깨짐 해결
plt.rc('font', family='Malgun Gothic')          # 맷플롯립 -부호 깨짐 해결





def createGraph(dfSum, param):
    '''dfSum, param'''
    if param == 'nomal':
        for i, col in enumerate(dfSum.columns):
            plt.subplot(len(dfSum.columns), 1, i+1)
            plt.plot(dfSum.index, dfSum[col], label=col)
            plt.grid()
            plt.legend()
        plt.show()
    
    elif param == 'ccr':
        for col in dfSum.columns:

            dpc = (dfSum[col] - dfSum[col].shift(1)) / dfSum[col].shift(1) * 100
            dpc.iloc[0] = 0
            cp = ((dpc+100)/100).cumprod()

            diff = (dfSum.index[-1] - dfSum.index[0]).days
            startBalance = float(dfSum.iloc[0][col])
            lastBalance = float(dfSum.iloc[-1][col])
            cagr = round(((lastBalance/startBalance)**(365/diff) - 1)*100, 2)

            profitRate = round((cp.iloc[-1]-1), 2)
            std = dpc.std()
            sharp = round(profitRate / std, 3)
            
            year_max = dfSum[col].rolling(252).max()
            drawdown = (dfSum[col] / year_max).rolling(252).min()
            mdd = str(round((drawdown.min() - 1)*100, 2)) + '%'

            label = '{}, cagr:{} , sharp:{}, mdd:{}'.format(col, cagr, sharp, mdd)
            plt.plot(dfSum.index, cp, label=label)
        
        plt.grid()
        plt.legend(loc='best')
        plt.title('수익률비교')
        plt.show()

    elif param == 'histogram':
        # plt.rc('font', family='Malgun Gothic')
        for i, col in enumerate(dfSum.columns):
            dpc = (dfSum[col]-dfSum[col].shift(1)) / dfSum[col].shift(1) * 100
            plt.subplot( len(dfSum.columns), 1, i+1)
            plt.hist(dpc, bins=100, label=col)
            plt.grid()
            plt.legend()
        plt.show()

    elif param == 'ma':
        for i, col in enumerate(dfSum.columns):
            ma5 = dfSum[col].rolling(5).mean()
            ma20 = dfSum[col].rolling(20).mean()
            ma60 = dfSum[col].rolling(60).mean()
            ma120 = dfSum[col].rolling(120).mean()
            ma250 = dfSum[col].rolling(250).mean()
            df = pd.DataFrame(
                {
                    'price':dfSum[col],
                    'ma5':ma5,
                    'ma20':ma20,
                    'ma60':ma60,
                    'ma120':ma120,
                    'ma250':ma250,
                }
            )
            plt.subplot(len(dfSum.columns), 1, i+1)
            plt.plot(df.index, df['price'], 'k-', label=col)
            plt.plot(df.index, df['ma5'], color='#2ca02c', linestyle='-', label='ma5')
            plt.plot(df.index, df['ma20'], color='#d62728', linestyle='-',  label='ma20')
            plt.plot(df.index, df['ma60'], color='#ff7f0e', linestyle='-', label='ma60')
            plt.plot(df.index, df['ma120'], color='#9467bd', linestyle='-', label='ma120')
            plt.plot(df.index, df['ma250'], color='#e377c2', linestyle='-', label='ma250')
            plt.grid()
            plt.legend(loc='best')
        plt.show()

    elif param == 'band':
        for i, col in enumerate(dfSum.columns):
            ma5 = dfSum[col].rolling(5).mean()
            ma20 = dfSum[col].rolling(20).mean()
            std = dfSum[col].rolling(20).std()
            upper = ma20 + 2*std
            lower = ma20 - 2*std
            df = pd.DataFrame(
                {
                    'price':dfSum[col],
                    'ma5':ma5,
                    'upper':upper,
                    'lower':lower
                }
            )
            plt.subplot(len(dfSum.columns), 1, i+1)
            plt.plot(df.index, df['price'], 'k-', label=col)
            plt.plot(df.index, df['upper'], 'k:')
            plt.plot(df.index, df['lower'], 'k:')
            plt.grid()
            plt.legend(loc='best')
        plt.show()





def createExcel(dfSum, dirPath, param):
    '''dfSum, dirPath, param'''
    if param == 'cor':
        writer = pd.ExcelWriter(dirPath)
        dfSum.to_excel(writer, sheet_name= '원자료')
        corr = dfSum.corr(method='pearson')
        corr.to_excel(writer, sheet_name= '상관계수')
        writer.save()
    print('엑셀 파일로 저장됨.')