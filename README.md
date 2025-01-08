# SmartManMailSender使用說明
更新日期: 2024-10-09 by Frank Huang
## 程式名稱
SmartManMailSender.exe

## 程式目的
一個可以透過Outlool/SMTP發送含附件的郵件給指定收件者的執行檔(exe)。

## 功能特色:
•	支援Outlook或是SMTP發送郵件。
•	支援HTML格式的信件內容 – 模板。
•	支援動態文字取代功能 – 預留文字。
•	支援附件檔案。
 

## 呼叫範例 (使用Outlook)
SmartManMailSender.exe outlook -sender frank@gmail.com -s "My title" -f "C:\\Desktop\\test1.pdf" -t "template.txt" -p placeholder1,placeholder2, "placeholder that has comma,"... 

## 參數說明 (使用Outlook)
SmartManMailSender.exe → 執行檔名稱
outlook → 寄送方式 (固定為outlook)
-sender frank@gmail.com → 寄件人員
-s "My title" → 信件標題 (若包含空白需使用雙引號包起來)
-f "C:\\Desktop\\test1.pdf" → 附件檔案路徑
-t "template.txt" → 信件樣板路徑
-p placeholder1,placeholder2, "placeholder that has comma,"...  → 預留文字內容，使用逗號隔開，若內容包含逗號整個預留文字需使用雙引號包起來





## 呼叫範例 (使用SMTP)
SmartManMailSender.exe smtp -host smtp.hibox.biz -port 587 -ssl false -username frank@smartman.com.tw -password mypassword -sender_name "志元資訊/Frank" -sender_email frank@smartman.com.tw -receiver_name "Receiver/Frank" -receiver_email frank@smartman.com.tw -s "My title" -f "test1.pdf" -t "template.txt" -p Frank,2024/10,frank@smartman.com.tw,"HR department",2024/10/10,"Octo 10, 2024","Frank Huang"

## 參數說明 (使用Outlook)
SmartManMailSender.exe → 執行檔名稱
smtp → 寄送方式，固定為smtp
                -host smtp.hibox.biz → SMTP伺服器
                -port 587
                -ssl false (非必要，預設為false)
                -username frank@smartman.com.tw → 登入帳號
                -password mypassword → 登入密碼
                -sender_name "志元資訊/Frank" → 寄件人名稱 (非必要)
                -sender_email frank@smartman.com.tw → 寄件人信箱
                -receiver_name "Receiver/Frank" → 收件人名稱 (非必要)
                -receiver_email frank@smartman.com.tw → 收件人信箱
                -s "My title" → 標題
                -f "test1.pdf" → 檔案路徑 (非必要) 
                -t "template.txt" → 樣板路徑
                -p Frank,2024/10,frank@smartman.com.tw,"HR  department",2024/10/10,"Octo 10, 2024","Frank Huang" → 預留文字，以逗號隔開 (非必要)

## 常見問題

1.	什麼是預留文字?
預留文字是在信件模板中，實際執行寄信程式時會取代預留文字的內容以置換成實際的內容。

2.	什麼是模板?
模板是在信件中所顯示的文字內容。模板中除了純文字的顯示，也可以透過鑲入預留文字(如[Placeholder1])與HTML格式已產生自訂樣式。設定方式請參考下方說明。

3.	可以不提供模板嗎?
不行，模板的內容為信件的文字內容，必須要提供。但是模板中不一定要提供預留文字。
4.	可以不提供附件檔案嗎?
可以，附件檔案為非必要參數。
5.	如何產生信件模板與設定預留文字? 
郵件模板以HTML格式的文本內容呈現，使用HTML格式可定義信件中文字的大小、字體、行距等樣式。前往 https://onlinehtmleditor.dev/ 或其他提供HTML編輯的網站，從中複製模板。以下為範例步驟：

(1)	前往https://onlinehtmleditor.dev/，透過上方工具欄中編輯下方的文字。
![image](https://github.com/user-attachments/assets/23bd7f6d-221d-4b8e-a7d2-e0b4fa4d858a)

 
(2)	設定預留文字:
模板可以任意數量的預留文字，名稱如[placeholder_1]、[placeholder_2]、...、[placeholder_n]。 [placeholder_n] 會被後來給的參數文字給取代，讓信件部份的文字可以置換成合適的資訊，如發薪年月、收件者名稱等文字。預留文字取代順序會與提供的預留文字參數順序相同。如參數提供-p Frank,2024/10,frank@smartman.com.tw,"HR department, Frank Huang"則[placeholder_1]=Frank、[placeholder_2]=2024/10、[placeholder_3]= frank@smartman.com.tw、[placeholder_4]= HR department, Frank Huang。

1.	留意預留文字的拼法placeholder，大小寫必須一致。
2.	預留文字可以重複使用。
3.	預留文字取代的順序會與執行程式給的參數順序相同。
(3)	完成編輯後，點選上方原始碼轉碼藍色按鈕。複製所有(HTML格式)文字再貼到新的文字檔並儲存，完成模板的新增。

該模板會是一個獨立的文字檔如template.txt，並可以儲存位於本機上任何位置，指定模板路徑時可使用絕對或相對路徑。

 範例template.txt如下:
 ```html
<pre>
<span style="font-size:14px"><span style="font-family:Arial,Helvetica,sans-serif">Dear <strong>[Placeholder1]</strong>,</span></span></pre>

<p><span style="font-size:22px"><strong>附件為您 <span style="font-family:Arial,Helvetica,sans-serif"><span style="color:#1abc9c">[Placeholder2]</span></span> 的薪資單</strong>。💰</span><br />
<span style="font-size:14px">若有問題歡迎聯繫[Placeholder3]。感謝您的付出與辛勞。</span>😁</p>

<p><span style="font-size:14px"><span style="font-family:Arial,Helvetica,sans-serif">We are pleased to provide you with your payslip for <span style="color:#1abc9c"><strong>[Placeholder2]</strong></span>.</span></span></p>

<p><span style="color:#2980b9"><span style="font-family:Arial,Helvetica,sans-serif">If you have any questions or concerns about your payslip, please do not hesitate to </span></span><span style="color:#2c3e50"><span style="font-family:Arial,Helvetica,sans-serif"><strong><span style="background-color:#f1c40f">contact our [Placeholder3]</span></strong></span></span><span style="color:#2980b9"><span style="font-family:Arial,Helvetica,sans-serif">.</span><br />
<span style="font-family:Arial,Helvetica,sans-serif">Thank you for your hard work and dedication. We appreciate your contributions to our organization.</span></span></p>

<p style="text-align:justify"><span style="font-size:10px">►發薪日為 [Placeholder5]。The payday is on [Placeholder6]</span></p>

<hr />
<p>[Placeholder7] 敬上。</p>

<p><span style="font-family:Arial,Helvetica,sans-serif"><span style="font-size:16px">Sincerely,</span><br />
<em>[Placeholder3]</em><br />
[Placeholder4]&trade;</span></p>

<p>&nbsp;</p>
```
6.	TD中如何呼叫? 範例如下：
Call SalLoadAppAndWait( 'SmartManMailSender.exe outlook -sender frank@gmail.com -s "My title" -f "C:\\Desktop\\test1.pdf" -t "template.txt" -p placeholder1,placeholder2, "placeholder that has comma,"', Window_NotVisible, nReturn )

7.	如何知道執行成功或是失敗?
在logs資料夾中文字檔如mail_log_20231025.txt可以查看執行歷程與錯誤訊息。
在TD中，若回傳0代表執行成功，非0代表失敗。

8.	出現錯誤Server execution failed (0x80080005 (CO_E_SERVER_EXEC_FAILURE))
方法一: Outlook需要取消自動發信功能，前往左上方File>左下Option>左邊Advanced>下滑到Send and receive>uncheck Send immediately when connected)，執行完後信件會儲存在寄件匣中。
 
方法二: 若本身Outlook有開啟自動發信功能，需要先關閉Outlook才得自動寄信。
