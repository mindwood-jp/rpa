#!/usr/bin/env python3
"""
在庫情報取得
"""
__author__  = "MindWood"
__version__ = "1.00"

from email import message
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import os
import sys
import time
import shutil
import pandas as pd
from skpy import Skype
import openpyxl as op

class MyClass:
	driver = None
	wait = None

	# ダウンロード先フォルダ
	download_path = '/home/mindwood/download'

	"""
	Webドライバに接続
	"""
	def attachment_driver(browser):
		if browser.lower() == 'firefox':
			options = webdriver.firefox.options.Options()
			if not sys.platform == 'win32': options.headless = True
			return webdriver.Firefox(options=options)
		elif browser.lower() == 'chrome':
			options = webdriver.ChromeOptions()
			if not sys.platform == 'win32': options.add_argument('--headless')
			options.add_experimental_option('detach', True)
			options.add_experimental_option('excludeSwitches', ['enable-logging'])
			options.add_experimental_option('prefs', {
				'download.default_directory': MyClass.download_path,
				'download.prompt_for_download': False,
				'download.directory_upgrade': True
			})
			options.use_chromium = True
			options.add_argument('start-maximized')
			options.add_argument('enable-automation')
			options.add_argument('--no-sandbox')
			options.add_argument('--disable-infobars')
			options.add_argument('--disable-extensions')
			options.add_argument('--disable-dev-shm-usage')
			options.add_argument('--disable-browser-side-navigation')
			options.add_argument('--disable-gpu')
			options.add_argument('--ignore-certificate-errors')
			options.add_argument('--ignore-ssl-errors')
			return webdriver.Chrome('/usr/local/bin/chromedriver', options=options)
		else:
			sys.exit(f'{browser} driver not found')

	"""
	Othello ログイン
	"""
	def othello_login():
		print('func: othello_login')

		MyClass.driver.set_window_size(950, 800)
		MyClass.driver.get('https://example.com')
		MyClass.wait = WebDriverWait(MyClass.driver, 10)
		MyClass.wait.until(EC.presence_of_all_elements_located)
		MyClass.driver.find_element(By.ID, '_user_id').send_keys('foo')  # ユーザID
		MyClass.driver.find_element(By.ID, '_user_pwd').send_keys('bar')  # パスワード
		MyClass.driver.find_element(By.ID, 'login_button').click()  # ログイン
		time.sleep(8)

	"""
	在庫情報ダウンロード
	"""
	def inventory_download():
		print('func: inventory_download')

		download_file = os.path.join(MyClass.download_path, 'Othello在庫データ.xlsx')
		# すでにあれば削除
		if os.path.isfile(download_file):
			os.remove(download_file)

		MyClass.driver.implicitly_wait(10)
		elmSidebarMenu = MyClass.driver.find_element(By.ID, 'sidebar-menu')
		elmSidebarMenu.find_element(By.XPATH, './/span[@data-id="MENU-L-5000"]').click()  # 在庫
		time.sleep(3)
		elmSidebarMenu.find_element(By.XPATH, './/span[@data-id="!INQZ0100I"]').click()  # 在庫照会
		time.sleep(5)
		elmPanel = MyClass.driver.find_element(By.CLASS_NAME, 'x_panel')
		elmPanel.find_element(By.XPATH, './/span[@data-id="*DOWNLOAD(ITEM)"]').click()  # ダウンロード(品目別)
		print('  Download link click')

		status = None
		while status != 'ダウンロードを完了しました。':
			time.sleep(10)
			status = elmPanel.find_element(By.XPATH, './/div[@id="output"]/span').text
			print('  Current status=' + status)
		
	"""
	在庫データ分析
	新しい品番が追加されたら、last.csvを削除すること
	廃番は気にしなくて良い
	"""
	def data_analysis():
		print('func: data_analysis')

		download_file = os.path.join(MyClass.download_path, 'Othello在庫データ.xlsx')
		while not os.path.isfile(download_file):
			time.sleep(3)

		MyClass.driver.quit()

		# pandas のバグ？回避
		if True:
			wb = op.load_workbook(download_file)
			wb.save(os.path.join(MyClass.download_path, '_Othello在庫データ.xlsx'))
			download_file = os.path.join(MyClass.download_path, '_Othello在庫データ.xlsx')

		current_file = os.path.join(MyClass.download_path, 'current.csv')
		last_file = os.path.join(MyClass.download_path, 'last.csv')
		message_file = os.path.join(MyClass.download_path, 'message.txt')

		# すでにあれば削除
		if os.path.isfile(message_file):
			os.remove(message_file)

		df = pd.read_excel(download_file, usecols=['品番', '品名', '現在庫数'], engine='openpyxl')
		df.to_csv(current_file, index=False, encoding='utf-8-sig')

		# 前回のファイルがあれば、それとの差分を表示
		if os.path.isfile(last_file):
			df2 = pd.read_csv(last_file)

			merged_df = pd.merge(df, df2, on='品番', suffixes=['_current', '_last'], how='outer')  # 外部結合
			merged_df['現在庫数_current'] = merged_df['現在庫数_current'].fillna(0).astype(int)
			merged_df['現在庫数_last'] = merged_df['現在庫数_last'].fillna(0).astype(int)
			merged_df = merged_df[merged_df['現在庫数_current'] != merged_df['現在庫数_last']]

			with open(message_file, mode='a', encoding='utf-8') as f:
				for _, row in merged_df.iterrows():
					name = row['品名_current'] if pd.isna(row['品名_last']) else row['品名_last']
					riseOrFall = '減りました' if row['現在庫数_last'] > row['現在庫数_current'] else '増えました'
					if row['現在庫数_current'] < 10:
						f.write(f"\n{row['品番']} {name} の在庫が {row['現在庫数_last']} から {row['現在庫数_current']} に{riseOrFall}")

		# 次回の差分抽出用に、前回ファイルを作成
		shutil.copyfile(current_file, last_file)

	"""
	在庫変動情報の転送
	"""
	def transfer():
		print('func: transfer')

		message_file = os.path.join(MyClass.download_path, 'message.txt')
		mes = None
		if os.path.isfile(message_file):
			with open(message_file, encoding='utf-8') as f:
				mes = ''.join(f.readlines()[1:101])

		# Skype 送信
		sk = Skype('bot@example.com', 'foobar')
		ch = sk.chats.chat('xxxxx@thread.skype')

		if not mes:
			ch.sendMsg('安全在庫が確保されています')
			return
		
		ch.sendMsg(mes)

		# paramiko

"""
メイン
"""
if __name__ == '__main__':
	MyClass.driver = MyClass.attachment_driver('Chrome')
	MyClass.othello_login()
	MyClass.inventory_download()
	MyClass.data_analysis()
	MyClass.transfer()
