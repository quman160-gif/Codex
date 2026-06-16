#!/bin/zsh

set -u

REPO_DIR="/Users/seroma/Desktop/簡宏志/codex"
EXCEL_FILE="班表codex/班表AI-11506-1拷貝.xlsx"
ICS_HOME="https://quman160-gif.github.io/Codex/"

print_subscription_urls() {
  echo "ICS 訂閱首頁："
  echo "$ICS_HOME"
  echo
  echo "全部人：https://quman160-gif.github.io/Codex/ics/all.ics"
  echo "佳惠：https://quman160-gif.github.io/Codex/ics/jiahui.ics"
  echo "士昂：https://quman160-gif.github.io/Codex/ics/shiang.ics"
  echo "宏志：https://quman160-gif.github.io/Codex/ics/hongzhi.ics"
  echo "恩婕：https://quman160-gif.github.io/Codex/ics/enjie.ics"
  echo "盈萱：https://quman160-gif.github.io/Codex/ics/yingxuan.ics"
  echo "重琦：https://quman160-gif.github.io/Codex/ics/zhongqi.ics"
  echo "鈺茜：https://quman160-gif.github.io/Codex/ics/yuxi.ics"
}

pause() {
  echo
  echo "按 Enter 關閉這個視窗..."
  read _
}

fail() {
  echo
  echo "失敗：$1"
  pause
  exit 1
}

echo "準備上傳班表 Excel 到 GitHub"
echo

cd "$REPO_DIR" || fail "找不到專案資料夾：$REPO_DIR"

command -v git >/dev/null 2>&1 || fail "找不到 git，請先安裝或確認 Xcode Command Line Tools。"

if [ ! -f "$EXCEL_FILE" ]; then
  fail "找不到 Excel 檔案：$EXCEL_FILE"
fi

echo "1/4 同步 GitHub 最新版本..."
git pull --rebase --autostash origin main || fail "git pull 失敗，請把上面的錯誤訊息截圖給 Codex。"

echo
echo "2/4 只加入指定 Excel 檔案..."
git add -- "$EXCEL_FILE" || fail "git add 失敗。"

if git diff --cached --quiet -- "$EXCEL_FILE"; then
  echo
  echo "Excel 沒有新的變更，不需要上傳。"
  echo
  print_subscription_urls
  pause
  exit 0
fi

COMMIT_TIME=$(date "+%Y-%m-%d %H:%M")
COMMIT_MESSAGE="Update schedule Excel $COMMIT_TIME"

echo
echo "3/4 建立 GitHub 更新紀錄..."
git commit -m "$COMMIT_MESSAGE" -- "$EXCEL_FILE" || fail "git commit 失敗。"

echo
echo "4/4 推送到 GitHub..."
git push || fail "git push 失敗，請把上面的錯誤訊息截圖給 Codex。"

echo
echo "完成！GitHub Actions 會自動重新產生 ICS。"
echo
print_subscription_urls
echo
echo "提醒：Google 日曆不一定立刻更新，通常要等 Google 自己重新抓取。"

pause
