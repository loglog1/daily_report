#!./env/Scripts/python
import win32com.client
import datetime
import pandas as pd
import os
from dotenv import load_dotenv
load_dotenv()

def mailbody():
    global myname
    today_task = pd.read_csv("./issues.csv",
                              usecols=[
                                '題名',
                                'プロジェクト', 
                                'ステータス', 
                                '期日',
                                '最新のコメント'  
                              ])
    projects = today_task['プロジェクト'].unique()

    # メール本文
    body = f"各位 \n\n \n\n おはようございます。{os.getenv("MY_NAME")}です。\n\n本日の日報を送付いたします。ご査収ください。\n\n \n\n"
    for pj_name in projects:
        body += pj_name + "------------------------------------------------------------------\n\n"
        for task_idx in today_task.loc[today_task["プロジェクト"]==pj_name,:].index:
            task = today_task.iloc[task_idx,:].to_dict()
            body += "【"+task['題名']+"】\n"
            # body += "プロジェクト："+task['プロジェクト'] +"\n"
            body += "状況：" + task['ステータス'] + "　期日：" + task['期日'] + "\n"
            body += "進捗："
            body += task['最新のコメント'] if type(task['最新のコメント']) is not float else "（題名のみ）"
            body += "\n\n \n\n"
    
    return body

def main():
    global myname
    # OutlookAPP のインスタンス化
    outlook = win32com.client.Dispatch("outlook.application")
    mapi = outlook.GetNamespace("MAPI")

    # メールオブジェクトの作成
    mail = outlook.CreateItem(0)  # 0: メールアイテム
    mail.bodyFormat = 2
    mail.To = os.getenv("TO_ADDRESS")
    mail.CC = os.getenv("CC_ADDRESS")
    today = datetime.datetime.today()
    mail.Subject = f"日報（{today.year}/{today.month}/{today.day}）{os.getenv("MY_NAME")}"
    mail.Body = mailbody()
    mail.Attachments.Add(os.getenv("AI_LIST"))

    mail.display()

if __name__ == "__main__":
    main()