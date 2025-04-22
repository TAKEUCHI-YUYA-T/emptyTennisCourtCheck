# -*- coding: utf-8 -*-
import smtplib
import email.message
import urllib.parse
import datetime
import settings
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait

option = Options()                          # オプションを用意
option.add_argument('--headless')           # ヘッドレスモードの設定を付与
#option.binary_location = "/"
driver = webdriver.Chrome(options=option)   # Chromeを準備(optionでヘッドレスモードにしている）
option.add_argument("--no-sandbox")
#driver = webdriver.Chrome()                #デバッグ時のみコメントアウト解除でバックグラウンド実行解除

#-------------------------設定ファイルにて定義---------------------
#Botチャンネルアクセストークン
CHANNEL_ACCESS_TOKEN = settings.CHANNEL_ACCESS_TOKEN
#BotユーザーID
USER_ID = settings.USER_ID
#対象URL
targetUrl = settings.targetUrl
#HTML保存ディレクトリ
saveFile = settings.saveFile
#検索期間（n日後まで）
serchDaysCount = settings.serchDaysCount
# 曜日のリスト
weekdays = settings.weekdays
#ボタン
xpMenu = settings.xpMenu
xpInquiry = settings.xpInquiry
xpNextPage = settings.xpNextPage
xpField = settings.xpField
xpPark = settings.xpPark
xpForward = settings.xpForward
#コート数
courtNum = settings.courtNum
#コートごとXPath要素
court1 = settings.court1
court2 = settings.court2
court3 = settings.court3
court4 = settings.court4
court5 = settings.court5
court6 = settings.court6
court7 = settings.court7
court8 = settings.court8
court9 = settings.court9
court10 = settings.court10
court11 = settings.court11
court12 = settings.court12
court13 = settings.court13
court14 = settings.court14
#送信メッセージ
template = settings.template
# 送信元と送信先のメールアドレス
senderEmail = settings.senderEmail
toAddr = settings.toAddr
password = settings.password
#----------------------------------------------------------------------
#現在日付と曜日を取得
today = datetime.date.today()
weekDay = datetime.date.today().strftime("%a")
intWeekDay = datetime.date.today().weekday()

def get_html(targetUrl):
	# URLを開く
	driver.get(targetUrl)

	#コート選択で使用する辞書を初期化
	courts = {}

	# elemの変数を用意して、elemの変数にクリックしたい「画像」のxpathの情報を格納し、クリック
	menuElem = driver.find_element(By.XPATH, xpMenu)
	menuElem.click()

	#URL遷移のパターンが異なることがあるので、「空き照会・予約の申込」ボタン押下はURL次第で判断
	url = driver.current_url
	parsedUrl = urllib.parse.urlparse(url)
	path = parsedUrl.path
	if path.endswith("Wg_KoukyouShisetsuYoyakuMoushikomi.aspx"):
		inquiryElem = driver.find_element(By.XPATH, xpInquiry)
		inquiryElem.click()

	nextPageElem = driver.find_element(By.XPATH, xpNextPage)
	nextPageElem.click()
	fieldElem = driver.find_element(By.XPATH, xpField)
	fieldElem.click()
	parkElem = driver.find_element(By.XPATH, xpPark)
	parkElem.click()
	forwardElem = driver.find_element(By.XPATH, xpForward)
	forwardElem.click()

	#曜日で絞り込み
	for i in range(serchDaysCount):
		#検索期間の曜日を計算
		futureWeekday = (intWeekDay + i + 1) % 7
		xpButton = '//*[@id="chk' + weekdays[futureWeekday] + '"]'
	    
	    # ボタンをクリック
		button = driver.find_element(By.XPATH, xpButton)
		button.click()

	forwardElem = driver.find_element(By.XPATH, xpForward)
	forwardElem.click()

	# 最大20秒間待機する
	wait = WebDriverWait(driver, 20)

	#テニスコートを選んで選択
	for i in range(1, courtNum + 1):
		court = globals()["court" + str(i)]  # globals()を使用して文字列から変数を取得
		courts["court" + str(i)] = court
	for court_key, court_value in courts.items():
		for j in range(serchDaysCount):
			date = today + datetime.timedelta(days=j + 1)
			formattedDate = date.strftime("%Y%m%d")
			xpCourt = (court_value + str(formattedDate) + '"]')
			targetCourt = driver.find_element(By.XPATH, xpCourt)
			targetCourt.click()


	forwardElem = driver.find_element(By.XPATH, xpForward)
	forwardElem.click()

	wait = WebDriverWait(driver, 20)

	pageHtml = driver.page_source

	# ブラウザを終了
	driver.quit()

	return pageHtml



def create_message(fromAddr, toAddr, subject, body):
    msg = email.message.Message()
    msg['Subject'] = subject
    msg['From'] = fromAddr
    msg['To'] = ",".join(toAddr)
    msg.add_header('Content-Type', 'text/html; charset=utf-8')
    body = body.encode('utf-8')
    msg.set_payload(body)
    return msg

def send(fromAddr, toAddr, msg):
    smtpobj = smtplib.SMTP('smtp.gmail.com', 587)
    smtpobj.starttls()
    smtpobj.login(senderEmail, password)
    smtpobj.sendmail(fromAddr, toAddr, msg.as_string().encode('utf-8'))
    smtpobj.close()
 
def main():
	subject = '【' + str(today) + '】墨田区営テニスコート 空き状況通知'
	pageHtml = get_html(targetUrl)
	body = template + pageHtml
	msg = create_message(senderEmail, toAddr, subject, body)
	send(senderEmail, toAddr, msg)

if __name__ == '__main__':
	try:
		main()
	except Exception as e:
		print("メールの送信中にエラーが発生しました:", str(e))