import pandas as pd
import os
import datetime as dt
os.chdir(r'M:\MSG Open Episodes\Apointments pull 24.01.22')

df = pd.read_csv('apptsFull.txt', '\t')

df.columns=['URN', 'EpisodeNumber', 'CareProvider', 'Speciality', 'Unknown', 'AppointmentDate', 'AppointmentTime', 'AptCareProvider', 'AptSpeciality', 'Group']
df['AppointmentDate'] =  pd.to_datetime(df['AppointmentDate'],errors='coerce')

df = df.sort_values(['URN','EpisodeNumber','AppointmentDate', 'AppointmentTime'],ascending=[True,False,False,False])
cols = list(df.columns)

ep_groups = df.groupby('EpisodeNumber')

results_list = []

for k,v in ep_groups:
    v['PreviousApt'] = v['AppointmentDate'].shift(-1)
    v['DaysSincePrivousAppiontment'] = v['AppointmentDate'] -  v['PreviousApt']
    df_out = v
    results_list.append(df_out)
    
df_res =  pd.concat(results_list)

df_res['DaysSincePrivousAppiontment'] = df_res['DaysSincePrivousAppiontment'].dt.total_seconds()

df_res['DaysSincePrivousAppiontment'] = df_res['DaysSincePrivousAppiontment'] / 60 / 60 / 24

df_res = df_res[df_res['DaysSincePrivousAppiontment'] > 365]

df_res.to_excel('Days between appointments.xlsx', index= False)