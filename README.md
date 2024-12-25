# FingerPrint

This is a webservice for produce Fingerprint reports.

python=3.10

mailmerge 蠻容易會有版本相容性的問題，所以要固定套件的版本。

## 2024-12-18 issues: 
> 1. 如果要使用 FastAPI 處理 Form 表單數據，需要 python-multipart 套件。(Form 來自 fastapi，而 python-multipart 是用來支持 multipart/form-data 的一個依賴庫。)
> 2. 如果要使用 FastAPI 則 index.html 中要使用 request.url_for 來取得靜態資源。並且 url_for 的參數要使用 filepath 而不是 filename。
> 3. 使用 FastAPI 時，request.url_for 會回傳 http 而不是 https，這導至無法成功取得靜態資源，例如:圖片與 css 檔案。(需要設定 nginx 反向代理才能解決，但嘗試很多次仍不會設定 zeabur)
> 4. 會想使用 FastAPI 是因為 Flask 的 request.url_for 是回傳 http，但後來發現可以設定 _scheme='https' 來解決。(但不確定為何 index.html 中使用 url_for 讀取靜態資源可以成功)