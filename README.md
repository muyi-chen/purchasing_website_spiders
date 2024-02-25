# 政府電子採購網爬蟲程式

本程式主要用於自動化爬取台灣政府電子採購網的招標資訊，便於企業快速查看並決定投標策略。

## 功能特點

- 根據指定的機關名稱、標案名稱、時間範圍自動爬取招標資訊。
- 自動保存爬取的資料至 `tender_info.xlsx` 文件，便於後續查閱。
- 支持設定定時執行，自動更新招標資訊。
- 避免重複記錄相同的標案編號，以免資料重複。

## 配置說明

程式的運行配置在 `config.ini` 文件中，以下是可配置的項目：

```ini
[DEFAULT]
org_name = 消防局             # 機關單位關鍵字
tender_name = 機器人           # 標案關鍵字
start_date = 2024/02/17        # 開始時間
end_date = 2024/02/23          # 結束時間
execution_interval = 1         # 執行間隔時間(小時)
```

## 使用方法

本程式提供兩種運行方式：

### 方式一：直接運行 `app.py`

1. 確保已安裝 Python 和所需的第三方庫（`requests`, `openpyxl`, `BeautifulSoup`等）。
    ```bash
    pip install requests beautifulsoup4 openpyxl
    ```
2. 將 `app.py` 和相關的 Python 模塊放置在同一目錄下。
3. 編輯 `config.ini` 文件，設置您的爬蟲參數。
4. 通過命令行運行 `app.py` 啟動程式：
    ```bash
    python app.py
    ```
5. 程式將根據設定定時自動執行爬蟲，結果保存在 `tender_info.xlsx`。

### 方式二：運行打包好的 `app.exe`

1. 確保 `app.exe`、`config.ini` 和 `tender_info.xlsx`（如果已存在）位於同一目錄下。
2. 編輯 `config.ini` 文件，設置您的爬蟲參數。
3. 直接雙擊 `app.exe` 文件或通過命令行運行：
    ```bash
    app.exe
    ```
4. 程式將根據設定定時自動執行爬蟲，結果保存在 `tender_info.xlsx`。這種方式不需要手動安裝 Python 或任何第三方庫。

## 注意事項

- 確保在運行程式前，`config.ini` 的配置正確無誤。
- 運行程式需要網路連接，以訪問政府電子採購網。
- 程式僅供學習和參考使用，請遵守相關法律法規，合法使用爬蟲技術。