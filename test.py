a = (10, 50)
b = (30, 40)
import numpy as np
print(np.asarray(a) - np.asarray(b))

# import pandas as pd
# import os
# from datetime import datetime
#
# print('Time started: ', datetime.now().strftime('%H:%M:%S'))
# b = False
#
# symbols = ['DEGOUSDT', 'DEXEUSDT', 'DFUSDT', 'DIAUSDT', 'DNTUSDT', 'DOCKUSDT', 'DOTDOWNUSDT', 'DOTUPUSDT', 'DREPUSDT', 'EDUUSDT', 'ELFUSDT', 'EOSBEARUSDT', 'EOSBULLUSDT', 'EOSDOWNUSDT', 'EOSUPUSDT', 'EPSUSDT', 'EPXUSDT', 'ERDUSDT', 'ERNUSDT', 'ETHBEARUSDT', 'ETHBULLUSDT', 'ETHDOWNUSDT', 'ETHUPUSDT', 'EURUSDT', 'FARMUSDT', 'FIDAUSDT', 'FILDOWNUSDT', 'FILUPUSDT', 'KDAUSDT', 'KEYUSDT', 'KMDUSDT', 'KP3RUSDT', 'LAZIOUSDT', 'LEVERUSDT', 'LINKDOWNUSDT', 'MDXUSDT', 'MFTUSDT', 'MIRUSDT', 'MITHUSDT', 'MLNUSDT', 'MOBUSDT', 'MOVRUSDT', 'MULTIUSDT', 'NANOUSDT', 'NBSUSDT', 'NEBLUSDT', 'NEXOUSDT', 'NMRUSDT', 'NPXSUSDT', 'NULSUSDT', 'OAXUSDT', 'OGUSDT', 'OMUSDT', 'ONGUSDT', 'OOKIUSDT', 'ORNUSDT', 'OSMOUSDT', 'OXTUSDT', 'PAXGUSDT', 'PAXUSDT', 'PEPEUSDT', 'PERLUSDT', 'PERPUSDT', 'PHAUSDT', 'PLAUSDT', 'PNTUSDT', 'POLSUSDT', 'POLYUSDT', 'POLYXUSDT', 'PONDUSDT', 'PORTOUSDT', 'POWRUSDT', 'PROMUSDT', 'PROSUSDT', 'PSGUSDT', 'PUNDIXUSDT', 'PYRUSDT', 'QIUSDT', 'QKCUSDT', 'QUICKUSDT', 'RADUSDT', 'RAMPUSDT', 'RAREUSDT', 'RDNTUSDT', 'REIUSDT', 'REPUSDT', 'REQUSDT', 'RGTUSDT', 'RIFUSDT', 'RPLUSDT', 'SANTOSUSDT', 'SCRTUSDT', 'SHIBUSDT', 'SLPUSDT', 'SNTUSDT', 'SSVUSDT', 'STEEMUSDT', 'STMXUSDT']
#
# for folder in symbols[60:90]:
#     try:
#         vals = pd.DataFrame()
#         try:
#             df = pd.read_csv(f'C:\\Users\\Abdykarim.D\\Documents\\dfs\\{folder}\\concatenated.csv')
#             print('Loaded concatenated')
#         except:
#             df = pd.DataFrame()
#             print('Could not load concatenated')
#         try:
#             df1 = pd.read_csv(f'C:\\Users\\Abdykarim.D\\Documents\\dfs\\{folder}\\feb-may.csv')
#             print('Loaded feb-may')
#         except:
#             df1 = pd.DataFrame()
#             print('Could not load feb-may')
#         df = pd.concat([df, df1])
#         df = df.reset_index()
#         df = df.drop(['index'], axis=1)
#         df = df.drop_duplicates(subset='Timestamp', keep='first')
#         df = df.reset_index()
#         print(f'Time started {folder}: ', datetime.now().strftime('%H:%M:%S'))
#         for i in range(15, len(df) - 1):
#             vals.loc[i, 'O'] = df['Open'].iloc[i]
#             vals.loc[i, 'H'] = df['High'].iloc[i]
#             vals.loc[i, 'L'] = df['Low'].iloc[i]
#             vals.loc[i, 'C'] = df['Close'].iloc[i]
#             for j in range(1, 16):
#                 vals.loc[i, f'O-{j}'] = df['Open'].iloc[i - j]
#                 vals.loc[i, f'H-{j}'] = df['High'].iloc[i - j]
#                 vals.loc[i, f'L-{j}'] = df['Low'].iloc[i - j]
#                 vals.loc[i, f'C-{j}'] = df['Close'].iloc[i - j]
#             diff_high = df['High'].iloc[i + 1] * 100.0 / vals['C'].iloc[i - 15] - 100
#             diff_low = df['Low'].iloc[i + 1] * 100.0 / vals['C'].iloc[i - 15] - 100
#     #         if i == 18:
#     #             print(df.iloc[i-15:i+5])
#     #             print(vals.iloc[i - 15])
#     #             print(diff_high, diff_low, abs(diff_high / diff_low), abs(diff_low / diff_high))
#     #             print(df['High'].iloc[i + 1], df['Low'].iloc[i + 1], '|', vals['C'].iloc[i - 15])
#     #             break
#             try:
#                 if diff_high != 0 and diff_low != 0 and diff_high >= 1 and abs(diff_high / diff_low) >= 2:
#                     vals.loc[i, 'Target'] = 1
#                 elif diff_high != 0 and diff_low != 0 and diff_low <= -1 and abs(diff_low / diff_high) >= 2:
#                     vals.loc[i, 'Target'] = -1
#                 else:
#                     vals.loc[i, 'Target'] = 0
#             except:
#                 vals.loc[i, 'Target'] = 0
#         vals.to_csv(f'C:\\Users\\Abdykarim.D\\Documents\\dfs\\{folder}\\concatenated_modified.csv', index=False)
#     except:
#         print('Error', folder)
# print('Time finished: ', datetime.now().strftime('%H:%M:%S'))
