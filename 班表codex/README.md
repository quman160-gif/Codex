# 班表日曆工具

這個資料夾會把 Excel 班表轉成可匯入日曆的 ICS / CSV。可以只在本機使用，不需要 GitHub、不需要 Google 授權，也不需要 Google API。

## 本機手動匯入模式

完成 Excel 班表後，雙擊：

```text
啟動本機日曆工具.command
```

或在 Terminal 執行：

```bash
cd /Users/seroma/Desktop/簡宏志/codex/班表codex
python3 -m pip install -r requirements.txt
python3 local_calendar_app.py
```

打開：

```text
http://127.0.0.1:8765/
```

在頁面選擇 Excel，按「產生檔案」，會輸出：

- 每個人的 `.ics`
- 每個人的 `.csv`
- 全部人的 `.ics`
- 全部人的 `.csv`

輸出位置在：

```text
local_exports/
```

Google 日曆網頁版手動匯入：右上角齒輪 → 設定 → 匯入與匯出 → 選擇 `.ics` 或 `.csv` 檔案 → 選日曆 → 匯入。

手動匯入是一次性動作；同一份班表重複匯入會產生重複事件。

## 產生 ICS

這是產生可部署訂閱網址的模式，不使用 GitHub 時可以略過。

```bash
cd /Users/seroma/Desktop/簡宏志/codex/班表codex
python3 -m pip install -r requirements.txt
python3 sync_schedule.py --base-url https://你的網域
```

輸出會在 `public/`：

- `public/ics/*.ics`：每個人的 ICS 訂閱檔
- `public/index.html`：每個人的訂閱網址列表
- `public/calendars.json`：機器可讀的訂閱清單
- `public/sync-report.json`：產生報告與未定義班別警告

## 班別規則

| 代碼 | 日曆顯示 | 時間 |
| --- | --- | --- |
| M | 早班 | 08:00-17:00 |
| M8 | 早班8-13 | 08:00-13:00 |
| M9 | 早班8.5-17.5 | 08:30-17:30 |
| M10 | 早班10-18 | 10:00-18:00 |
| M15 | 早班8-15 | 08:00-15:00 |
| M16 | 早班8-16 | 08:00-16:00 |
| E | 晚班 | 17:00-22:00 |
| E13 | 晚班13-22 | 13:00-22:00 |
| E14 | 晚班14-22 | 14:00-22:00 |
| E15 | 晚班15-22 | 15:00-22:00 |
| E16 | 晚班16-22 | 16:00-22:00 |
| A | 整天 | 12:00-22:00 |
| S | 二頭班 8-13，17-21 | 產生兩個活動 |
| O | 休息 | 整天活動，ICS 內加上紅色 `COLOR:#D50000` |

Excel 目前也有 `早`、`午`、`晚`、`8.5 17` 這類兼職班別，程式會一起轉成活動。`x` 和空白會略過。

## Google 日曆訂閱

把 `public/` 放到任何穩定的 HTTPS 靜態網站後，每個人的網址會像：

```text
https://你的網域/ics/hongzhi.ics
```

Google 日曆網頁版：左側「其他日曆」旁的 `+` →「從網址」→ 貼上 `.ics` 訂閱網址。

iPhone：設定 → 日曆 → 帳號 → 加入帳號 → 其他 → 加入訂閱的行事曆。

Android：先在 Google 日曆網頁版訂閱，手機上的 Google 日曆 App 會同步顯示。

Google 會自行抓取 ICS 更新，但重新抓取不是即時的；只要 HTTPS 上的 `.ics` 檔案被更新，之後就會自動反映。

## GitHub Pages 自動更新

這是可選方案；只做本機手動匯入時可以略過。

專案根目錄已放好 `.github/workflows/build-schedule-ics.yml`。把整個專案推到 GitHub 後，到 GitHub repo 的 Settings → Pages，把 Build and deployment 設為 GitHub Actions。

之後只要更新 `班表codex` 裡的 Excel 並推上 GitHub，Actions 會自動：

1. 安裝 `openpyxl`
2. 讀取 Excel 班表
3. 重建 `public/ics/*.ics`
4. 部署到 GitHub Pages

預設訂閱網址會是：

```text
https://你的GitHub帳號.github.io/你的repo名稱/ics/hongzhi.ics
```

如果你使用自訂網域，請把 workflow 裡 `--base-url` 後面的網址換成你的正式網址。

## 顏色說明

`休息` 已在 ICS 事件裡寫入 `COLOR:#D50000`。不同日曆 App 對 ICS 單一事件顏色的支援不完全一致；iPhone 日曆通常較能讀取 ICS 額外屬性，Google 日曆若沒有顯示紅色，資料本身仍會是整天休息事件。
