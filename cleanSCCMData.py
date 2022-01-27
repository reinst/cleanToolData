import os
import shutil
import datetime
import pandas as pd

today_time = today_time = datetime.datetime.now().strftime('%d-%m-%Y-%S')
def main(tool):
    move_files_to_share_folder(tool)
    clean_base_data(tool)
    find_missing_agents(tool)
    agent_percent_install_rate(tool)

def move_src_data(tool):
    original = [fr'\\server\{tool}\Missing Security Agent - {tool}.xls',
                fr'\\server\{tool}\{tool} Versions.xls']
    target = fr'\\server\{tool}\src_data'
    for f in original:
        shutil.move(f,target)
    print('Done')


def timestamp_file(tool):
    os.rename(fr'\\server\{tool}\src_data\Missing Security Agent - {tool}.xls',
                fr'\\server\{tool}\src_data\{today_time}_Missing Security Agent - {tool}.xls')  
    os.rename(fr'\\server\{tool}\src_data\{tool} Versions.xls',
                fr'\\server\{tool}\src_data\{today_time}_{tool} Versions.xls')
    print('Done')

def move_files_to_share_folder(tool):
    original = [fr'C:\Users\username\Desktop\Missing Security Agent - {tool}.xls',
                fr'C:\Users\username\Desktop\{tool} Versions.xls']
    target = fr'\\server\{tool}'
    for f in original:
        shutil.copy(f,target)
    print('Done')

def clean_base_data(tool):
    agent01 = pd.read_excel(fr'\\server\{tool}\\{tool} Versions.xls', skiprows=6)
    agent02 = agent01.rename(columns={'File Version':'Agent Version'})
    agent02.drop(['File Name'], axis=1, inplace=True)
    agent02.dropna(how='all', axis=1, inplace=True)
    agent02['AD Site Name'] = agent02['AD Site Name'].str.split('-').str[:2].str.join('-')
    agent02['Agent Version'] = agent02['Agent Version'].str.replace(', ', '.')
    agent02.to_csv(fr'\\server\{tool}\{tool}_AgentVersion.csv', index=False)
    print('Done')

def find_missing_agents(tool):
    missing01 = pd.read_excel(fr'\\server\{tool}\\Missing Security Agent - {tool}.xls', skiprows=6)
    missing02 = missing01.rename(columns={'File Version':'Agent Version'})
    missing02.drop(['OS Name', 'OS Build'], axis=1, inplace=True)
    missing02.dropna(how='all', axis=1, inplace=True)
    missing02['AD Site Name'] = missing02['AD Site Name'].str.split('-').str[:2].str.join('-')
    missing02['Agent Version'] = 'No Agent'
    missing02.to_csv(fr'\\server\{tool}\{tool}_AgentVersion.csv', mode='a', header=False, index=False)
    print('Done')

def agent_percent_install_rate(tool):
    agent = pd.read_csv(fr'\\server\{tool}\{tool}_AgentVersion.csv', usecols =['Agent Version'])
    agent_value = agent.groupby(['Agent Version'])['Agent Version'].count().rename('Percentage').transform(lambda x: x/x.sum()*100).round(2)
    agent_value.to_csv(fr'\\server\{tool}\{tool}_AgentInstallRate.csv')
    print('Done')


tools = ['Forcepoint']
for i in tools:
    try:
        if __name__ == '__main__':
            main(i)
    except:
        print(f"Tool: {i} is having issues processing data")
        continue
