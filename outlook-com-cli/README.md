# Classic Outlook COM 自動草稿工具

這版使用本機 Classic Outlook COM，不走 Microsoft Graph，不需要 Client ID。

## 前置條件

- Windows
- Classic Outlook 桌面版已開啟並登入信箱
- Python 安裝 `pywin32`

安裝：

```powershell
python -m pip install pywin32
```

## 預覽

```powershell
cd C:\Users\arthurhung\Desktop\workspace\outlook-com-cli
python outlook_com_cli.py --dry-run
```

或雙擊：

```text
預覽今日已完成草稿.bat
```

## 建立草稿

```powershell
python outlook_com_cli.py
```

或雙擊：

```text
建立今日已完成草稿.bat
```

它會：

1. 從目前 Outlook 預設收件匣找今天主旨包含 `(已完成)` 的最新信。
2. 找今天主旨包含 `ETL_TW` 的最新信。
3. 從第二封主旨抽出 `hh:mm:ss`。
4. 對第一封建立全部回覆草稿。
5. 修改草稿主旨。
6. 在原信表格最後新增「網銀資料 / 完成 / 完成時間」列。
7. 儲存並開啟全部回覆草稿視窗，但不寄出。

如果只想建立草稿、不開啟 Classic Outlook 的獨立草稿視窗，可加上：

```powershell
python outlook_com_cli.py --no-display
```

## 指定信箱

如果 Outlook 有多個信箱，可指定 SMTP：

```powershell
python outlook_com_cli.py --mailbox 00544984@cathaybk.com.tw
```
