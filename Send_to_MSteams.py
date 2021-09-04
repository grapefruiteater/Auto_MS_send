
from socket import AddressInfo
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.alert import Alert

import os
import glob
import shutil
import time

import configparser

# ログイン情報をconfig.iniファイルから取得
inifile = configparser.ConfigParser()
inifile.read(r"C:/SeleConf/configure/config.ini", 'UTF-8') ###################
user_email = inifile.get('user', 'user_email')
try:
    user_pw = inifile.get('user', 'user_pw')
except:
    pass
userid = user_email.split('@')[0]

Address = ["username"]
DIDs = ["a"]

#for Name in Address:
for i,Name in enumerate(Address):
    # カレントディレクトリの移動
    os.chdir('C:/Users/' + userid + '/Downloads')

    # カレントディレクトリの取得
    current_dir = os.getcwd()

    # ローカルにtmpDownloadフォルダを設定
    tmp_download_dir = f'{current_dir}\\tmpDownload'
    try:
        shutil.rmtree(tmp_download_dir)
    except:
        print('exist') 


    # Chromeオプション設定でダウンロード先をtmpDownloadフォルダに変更 ・・・ ★
    options = webdriver.ChromeOptions()
    prefs = {'download.default_directory' : tmp_download_dir }
    options.add_experimental_option('prefs',prefs)

    # ドライバのパス設定, chrome用jのwebdriverはあらかじめダウンロードが必要
    driver_path = r'//chromedriver.exe'

    # ★で設定したオプションを適用してChromeを起動
    driver = webdriver.Chrome(executable_path = driver_path, chrome_options = options)

    # Web版 Microsoft teamsへの画面遷移
    #driver.get('https://teams.microsoft.com/l/chat/0/0?users=' + DID +'&msLaunch=false&message=')
    driver.get('https://teams.microsoft.com/dl/launcher/launcher.html?url=%2F_%23%2Fl%2Fchat%2F0%2F0%3Fusers%3D' + Name +'%40micron.com%26msLaunch%3Dfalse%26message%3D&type=chat&deeplinkId=c4c91dc3-0837-449d-af7b-5abdf91a33c9&directDl=true&msLaunch=false&enableMobilePage=true&suppressPrompt=true')

    # フルスクリーン, 画面が小さいと表示されないhtml_idにはアクセスできなくなるため。
    driver.maximize_window()

    # 画面遷移を10秒待機
    time.sleep(3)
    try:
        alert = driver.switch_to.alert
        text = alert.text
        alert.dismiss()
    except:
        pass

    btn = driver.find_element_by_xpath("/html/body/div/div/div/div[2]/div/button[2]")
    btn.click()
    time.sleep(5)

    # ログイン処理 ： メールアドレスを入力
    try:
        driver.find_element_by_id('i0116').send_keys(user_email)
        driver.find_element_by_id('idSIButton9').click()
    except:
        print('Error')

    #　画面遷移を待機
    time.sleep(25)

    YIP_file_path="/Latest_YIP_Upload.csv"

    timediff = 0
    try:
        nowt = time.time()
        filet = os.path.getmtime(YIP_file_path)
        timediff = nowt - filet
    except:
        pass

    if os.path.exists(YIP_file_path) and timediff < 3*60*60*24:
        print("Exist")
    else:
        driver.find_element_by_xpath("/html/body/div[2]/div[2]/div[2]/div[1]/div/messages-header/div[2]/div/message-pane/div/div[3]/new-message/div/div[2]/form/div[3]/div[1]/div[2]/div/div/div[2]").click()
        driver.find_element_by_xpath("/html/body/div[2]/div[2]/div[2]/div[1]/div/messages-header/div[2]/div/message-pane/div/div[3]/new-message/div/div[2]/form/div[3]/div[1]/div[2]/div/div/div[2]/div/div/div/div").send_keys(Keys.BACK_SPACE)
        driver.find_element_by_xpath("/html/body/div[2]/div[2]/div[2]/div[1]/div/messages-header/div[2]/div/message-pane/div/div[3]/new-message/div/div[2]/form/div[3]/div[1]/div[2]/div/div/div[2]/div/div/div/div").send_keys('%s-san %s YIP upload file is nothing or old. \n Please update bellow csv file. \\\\WINNTDOM\\root\\Common\\MMJ\\Secure\\PIE\\All-Write\\YE\\Report\\HVM\\Weekly_material\\120s\\YIP\\Weekly_Dieloss_Pareto\\YIP_up_csv\\%s_Latest_YIP_Upload.csv'%(Name,DIDs[i],DIDs[i]))
        driver.find_element_by_id('send-message-button').click()
        print("Nothing")

    # 念のための待機時間
    time.sleep(2)

    # closeする
    driver.quit()

