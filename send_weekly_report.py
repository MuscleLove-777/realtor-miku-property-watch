#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
send_weekly_report.py
Googleスプレッドシートから直近7日間のデータを取得し、
週次HTMLメールレポートを生成してGmailの下書きに保存する。

認証:
  - スプレッドシート: 環境変数 GOOGLE_CREDENTIALS_JSON（サービスアカウントJSON文字列）
  - Gmail: 環境変数 GMAIL_TOKEN_JSON（OAuthトークン）またはサービスアカウント委任
  - Gmail認証失敗時はHTMLファイルをローカルに保存（フォールバック）
"""

import os
import sys
import json
import base64
import logging
from datetime import datetime, date, timedelta
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

import gspread
from google.oauth2.service_account import Credentials as SACredentials
from google.oauth2.credentials import Credentials as OAuthCredentials
from googleapiclient.discovery import build

# ───────────────────────────── ログ設定 ─────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
logger = logging.getLogger(__name__)

# ───────────────────────────── 定数 ─────────────────────────────
SPREADSHEET_NAME = "不動産ウォッチ_ハイハイタウン"

SCOPES_SHEETS = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]
SCOPES_GMAIL = [
    "https://www.googleapis.com/auth/gmail.compose",
]


# ───────────────────────────── 認証ヘルパー ─────────────────────────────
def get_sheets_client() -> gspread.Client:
    """スプレッドシート用のgspreadクライアントを返す。"""
    creds_json_str = os.environ.get("GOOGLE_CREDENTIALS_JSON")
    if not creds_json_str:
        raise EnvironmentError(
            "環境変数 GOOGLE_CREDENTIALS_JSON が設定されていません。"
        )
    creds_info = json.loads(creds_json_str)
    credentials = SACredentials.from_service_account_info(
        creds_info, scopes=SCOPES_SHEETS
    )
    return gspread.authorize(credentials)


def get_gmail_service():
    """Gmail APIサービスを返す。認証失敗時はNoneを返す。"""
    # 方法1: OAuthトークンJSON
    token_json_str = os.environ.get("GMAIL_TOKEN_JSON")
    if token_json_str:
        try:
            token_info = json.loads(token_json_str)
            credentials = OAuthCredentials.from_authorized_user_info(
                token_info, scopes=SCOPES_GMAIL
            )
            service = build("gmail", "v1", credentials=credentials)
            logger.info("Gmail API: OAuthトークンで認証成功")
            return service
        except Exception as e:
            logger.warning(f"Gmail OAuthトークン認証失敗: {e}")

    # 方法2: サービスアカウント（ドメイン全体の委任）
    creds_json_str = os.environ.get("GOOGLE_CREDENTIALS_JSON")
    if creds_json_str:
        try:
            creds_info = json.loads(creds_json_str)
            credentials = SACredentials.from_service_account_info(
                creds_info,
                scopes=SCOPES_GMAIL,
            )
            # 委任ユーザーが設定されている場合
            delegate_email = os.environ.get("GMAIL_DELEGATE_EMAIL")
            if delegate_email:
                credentials = credentials.with_subject(delegate_email)

            service = build("gmail", "v1", credentials=credentials)
            logger.info("Gmail API: サービスアカウント委任で認証成功")
            return service
        except Exception as e:
            logger.warning(f"Gmail サービスアカウント認証失敗: {e}")

    logger.warning("Gmail認証に失敗。ローカルファイルへフォールバックします。")
    return None


# ───────────────────────────── スプレッドシート読み込み ─────────────────────────────
def fetch_sheet_data(
    client: gspread.Client,
    sheet_name: str,
    days: int = 7,
) -> list[dict]:
    """指定シートから直近N日間のデータを辞書リストで返す。"""
    spreadsheet = client.open(SPREADSHEET_NAME)
    worksheet = spreadsheet.worksheet(sheet_name)
    all_records = worksheet.get_all_records()

    cutoff = (date.today() - timedelta(days=days)).isoformat()
    filtered = [r for r in all_records if r.get("日付", "") >= cutoff]
    logger.info(f"{sheet_name}: 直近{days}日間で {len(filtered)}件取得")
    return filtered


def fetch_summary_data(
    client: gspread.Client,
    weeks: int = 4,
) -> list[dict]:
    """価格推移サマリーから直近N週間分のデータを返す。"""
    spreadsheet = client.open(SPREADSHEET_NAME)
    worksheet = spreadsheet.worksheet("価格推移サマリー")
    all_records = worksheet.get_all_records()

    cutoff = (date.today() - timedelta(weeks=weeks)).isoformat()
    filtered = [r for r in all_records if r.get("日付", "") >= cutoff]
    logger.info(f"価格推移サマリー: 直近{weeks}週間で {len(filtered)}件取得")
    return filtered


# ───────────────────────────── 変動検出 ─────────────────────────────
def detect_changes(
    current: list[dict],
    client: gspread.Client,
) -> dict:
    """
    新規売出し・掲載終了・価格変更を検出する。
    直近7日以内のデータと7〜14日前のデータを比較する。
    """
    prev = fetch_sheet_data(client, "ハイハイタウン売出し", days=14)
    cutoff_7 = (date.today() - timedelta(days=7)).isoformat()

    # 7日より前のデータ
    older = [r for r in prev if r.get("日付", "") < cutoff_7]
    # 現在のURL集合と過去のURL集合
    current_urls = {r.get("URL", "") for r in current if r.get("URL")}
    older_urls = {r.get("URL", "") for r in older if r.get("URL")}

    # 新規売出し: 現在にあって過去にない
    new_listings = [r for r in current if r.get("URL", "") not in older_urls]

    # 掲載終了: 過去にあって現在にない
    removed_listings = [r for r in older if r.get("URL", "") not in current_urls]

    # 価格変更: 同一URLで価格が異なる
    older_price_map = {}
    for r in older:
        url = r.get("URL", "")
        if url:
            older_price_map[url] = r.get("価格(万円)", 0)

    price_changes = []
    for r in current:
        url = r.get("URL", "")
        if url in older_price_map:
            old_price = older_price_map[url]
            new_price = r.get("価格(万円)", 0)
            if old_price and new_price and old_price != new_price:
                price_changes.append({
                    "物件名": r.get("物件名", ""),
                    "旧価格": old_price,
                    "新価格": new_price,
                    "変動": new_price - old_price,
                    "URL": url,
                })

    return {
        "new": new_listings,
        "removed": removed_listings,
        "price_changes": price_changes,
    }


# ───────────────────────────── HTML生成 ─────────────────────────────
def generate_html_report(
    haihai: list[dict],
    nearby: list[dict],
    summary: list[dict],
    changes: dict,
    report_date: str,
) -> str:
    """週次レポートのHTML文字列を生成する。"""

    # ─── インラインCSSスタイル ───
    style = """
    <style>
        body { font-family: 'Hiragino Sans', 'Yu Gothic', 'Meiryo', sans-serif; color: #333; margin: 0; padding: 20px; background: #f5f5f5; }
        .container { max-width: 800px; margin: 0 auto; background: #fff; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); padding: 30px; }
        h1 { color: #2c3e50; border-bottom: 3px solid #3498db; padding-bottom: 10px; font-size: 20px; }
        h2 { color: #2c3e50; border-left: 4px solid #3498db; padding-left: 12px; font-size: 16px; margin-top: 30px; }
        h3 { color: #555; font-size: 14px; }
        table { border-collapse: collapse; width: 100%; margin: 15px 0; font-size: 13px; }
        th { background: #3498db; color: #fff; padding: 8px 10px; text-align: left; white-space: nowrap; }
        td { padding: 7px 10px; border-bottom: 1px solid #eee; }
        tr:nth-child(even) { background: #f9f9f9; }
        tr:hover { background: #eef5ff; }
        .price-up { color: #e74c3c; font-weight: bold; }
        .price-down { color: #2980b9; font-weight: bold; }
        .new-badge { background: #e74c3c; color: #fff; padding: 2px 6px; border-radius: 3px; font-size: 11px; }
        .removed-badge { background: #95a5a6; color: #fff; padding: 2px 6px; border-radius: 3px; font-size: 11px; }
        .summary-box { background: #eaf2f8; border-radius: 6px; padding: 15px; margin: 15px 0; }
        .comment-box { background: #fef9e7; border: 1px dashed #f0c040; border-radius: 6px; padding: 15px; margin: 20px 0; }
        .footer { text-align: center; color: #888; font-size: 12px; margin-top: 30px; padding-top: 15px; border-top: 1px solid #ddd; }
        a { color: #3498db; text-decoration: none; }
        a:hover { text-decoration: underline; }
    </style>
    """

    # ─── ヘッダー ───
    html = f"""<!DOCTYPE html>
<html lang="ja">
<head><meta charset="UTF-8">{style}</head>
<body>
<div class="container">
<h1>上本町ハイハイタウン 市場動向レポート</h1>
<p>レポート期間: {report_date} 週</p>
"""

    # ─── セクション1: ハイハイタウン現在の売出し物件一覧 ───
    html += '<h2>ハイハイタウン 現在の売出し物件一覧</h2>'
    if haihai:
        html += """<table>
<tr><th>物件名</th><th>価格(万円)</th><th>面積(㎡)</th><th>㎡単価</th><th>階数</th><th>間取り</th><th>詳細</th></tr>"""
        for r in haihai:
            url = r.get("URL", "")
            link = f'<a href="{url}">詳細</a>' if url else "-"
            html += f"""<tr>
<td>{r.get('物件名', '-')}</td>
<td>{r.get('価格(万円)', '-')}</td>
<td>{r.get('面積(㎡)', '-')}</td>
<td>{r.get('㎡単価(万円)', '-')}</td>
<td>{r.get('階数', '-')}</td>
<td>{r.get('間取り', '-')}</td>
<td>{link}</td>
</tr>"""
        html += '</table>'
    else:
        html += '<p>現在売出し中の物件はありません。</p>'

    # ─── セクション2: 新規売出し・価格変更・掲載終了 ───
    html += '<h2>今週の変動</h2>'

    # 新規売出し
    if changes["new"]:
        html += '<h3><span class="new-badge">NEW</span> 新規売出し</h3><table>'
        html += '<tr><th>物件名</th><th>価格(万円)</th><th>面積(㎡)</th><th>間取り</th></tr>'
        for r in changes["new"]:
            html += f"""<tr>
<td>{r.get('物件名', '-')}</td>
<td>{r.get('価格(万円)', '-')}</td>
<td>{r.get('面積(㎡)', '-')}</td>
<td>{r.get('間取り', '-')}</td>
</tr>"""
        html += '</table>'

    # 価格変更
    if changes["price_changes"]:
        html += '<h3>価格変更</h3><table>'
        html += '<tr><th>物件名</th><th>旧価格</th><th>新価格</th><th>変動額</th></tr>'
        for c in changes["price_changes"]:
            diff = c["変動"]
            css_class = "price-up" if diff > 0 else "price-down"
            sign = "+" if diff > 0 else ""
            html += f"""<tr>
<td>{c['物件名']}</td>
<td>{c['旧価格']}万円</td>
<td>{c['新価格']}万円</td>
<td class="{css_class}">{sign}{diff}万円</td>
</tr>"""
        html += '</table>'

    # 掲載終了
    if changes["removed"]:
        html += '<h3><span class="removed-badge">終了</span> 掲載終了</h3><table>'
        html += '<tr><th>物件名</th><th>価格(万円)</th><th>面積(㎡)</th></tr>'
        for r in changes["removed"]:
            html += f"""<tr>
<td>{r.get('物件名', '-')}</td>
<td>{r.get('価格(万円)', '-')}</td>
<td>{r.get('面積(㎡)', '-')}</td>
</tr>"""
        html += '</table>'

    if not changes["new"] and not changes["price_changes"] and not changes["removed"]:
        html += '<p>今週は変動なしです。</p>'

    # ─── セクション3: 周辺相場との比較 ───
    html += '<h2>周辺相場との比較</h2>'
    if nearby:
        html += """<table>
<tr><th>物件名</th><th>最寄駅</th><th>価格(万円)</th><th>面積(㎡)</th><th>㎡単価</th><th>築年数</th></tr>"""
        for r in nearby:
            html += f"""<tr>
<td>{r.get('物件名', '-')}</td>
<td>{r.get('最寄駅', '-')}</td>
<td>{r.get('価格(万円)', '-')}</td>
<td>{r.get('面積(㎡)', '-')}</td>
<td>{r.get('㎡単価(万円)', '-')}</td>
<td>{r.get('築年数', '-')}</td>
</tr>"""
        html += '</table>'

    # 周辺平均 vs HT平均の比較サマリー
    if summary:
        latest = summary[-1]
        ht_avg = latest.get("HT平均㎡単価", "-")
        nb_avg = latest.get("周辺平均㎡単価", "-")
        html += f"""<div class="summary-box">
<strong>最新㎡単価比較:</strong><br>
ハイハイタウン平均: <strong>{ht_avg}万円/㎡</strong> ｜
周辺平均: <strong>{nb_avg}万円/㎡</strong>
</div>"""

    # ─── セクション4: 直近4週の㎡単価推移 ───
    html += '<h2>直近4週の㎡単価推移</h2>'
    if summary:
        html += """<table>
<tr><th>日付</th><th>HT平均㎡単価</th><th>周辺平均㎡単価</th><th>HT売出件数</th><th>周辺売出件数</th></tr>"""
        for s in summary:
            html += f"""<tr>
<td>{s.get('日付', '-')}</td>
<td>{s.get('HT平均㎡単価', '-')}</td>
<td>{s.get('周辺平均㎡単価', '-')}</td>
<td>{s.get('HT売出件数', '-')}</td>
<td>{s.get('周辺売出件数', '-')}</td>
</tr>"""
        html += '</table>'
    else:
        html += '<p>推移データがまだ蓄積されていません。</p>'

    # ─── 美玖のコメント欄 ───
    html += """<h2>美玖のコメント</h2>
<div class="comment-box">
<p>【ここに美玖のコメントを記入】</p>
<p>今週の市場動向について、気になるポイントや投資判断に関するアドバイスを記載してください。</p>
</div>"""

    # ─── フッター ───
    html += f"""
<div class="footer">
<p>不動産コンサルタント 美玖</p>
<p>本レポートは自動生成されています。データ出典: SUUMO</p>
<p>生成日時: {datetime.now().strftime('%Y/%m/%d %H:%M')}</p>
</div>
</div>
</body>
</html>"""

    return html


# ───────────────────────────── Gmail下書き作成 ─────────────────────────────
def create_gmail_draft(service, html_content: str, subject: str, to_email: str = "me"):
    """Gmail APIで下書きを作成する。"""
    message = MIMEMultipart("alternative")
    message["Subject"] = subject
    message["To"] = to_email
    message["From"] = "me"

    # HTMLパート
    html_part = MIMEText(html_content, "html", "utf-8")
    message.attach(html_part)

    # Base64エンコード
    raw = base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")

    draft = service.users().drafts().create(
        userId="me",
        body={"message": {"raw": raw}},
    ).execute()

    logger.info(f"Gmail下書き作成成功: Draft ID = {draft['id']}")
    return draft


def save_html_fallback(html_content: str, report_date: str) -> str:
    """Gmail認証失敗時にHTMLをローカルファイルに保存する。"""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    filename = f"weekly_report_{report_date.replace('/', '-')}.html"
    filepath = os.path.join(script_dir, filename)

    with open(filepath, "w", encoding="utf-8") as f:
        f.write(html_content)

    logger.info(f"HTMLレポートをローカルに保存: {filepath}")
    return filepath


# ───────────────────────────── メイン処理 ─────────────────────────────
def main():
    """メイン処理: スプレッドシートからデータ取得→レポート生成→Gmail下書き作成。"""
    today = date.today()
    report_date = today.strftime("%Y/%m/%d")
    subject = f"【美玖・不動産レポート】{report_date}週 上本町ハイハイタウン市場動向"

    # ─── スプレッドシートからデータ取得 ───
    try:
        sheets_client = get_sheets_client()
    except Exception as e:
        logger.error(f"スプレッドシート認証エラー: {e}")
        sys.exit(1)

    try:
        haihai = fetch_sheet_data(sheets_client, "ハイハイタウン売出し", days=7)
        nearby = fetch_sheet_data(sheets_client, "周辺相場", days=7)
        summary = fetch_summary_data(sheets_client, weeks=4)
    except Exception as e:
        logger.error(f"スプレッドシート読み込みエラー: {e}")
        sys.exit(1)

    # ─── 変動検出 ───
    try:
        changes = detect_changes(haihai, sheets_client)
    except Exception as e:
        logger.warning(f"変動検出でエラー（スキップ）: {e}")
        changes = {"new": [], "removed": [], "price_changes": []}

    # ─── HTMLレポート生成 ───
    html_content = generate_html_report(
        haihai, nearby, summary, changes, report_date
    )
    logger.info(f"HTMLレポート生成完了 ({len(html_content)} bytes)")

    # ─── Gmail下書き作成（またはローカル保存） ───
    gmail_service = get_gmail_service()

    if gmail_service:
        try:
            # 宛先メールアドレス（環境変数で設定可能）
            to_email = os.environ.get("REPORT_TO_EMAIL", "me")
            draft = create_gmail_draft(gmail_service, html_content, subject, to_email)
            logger.info("週次レポートをGmail下書きに保存しました。")
        except Exception as e:
            logger.error(f"Gmail下書き作成失敗: {e}")
            filepath = save_html_fallback(html_content, report_date)
            logger.info(f"フォールバック: {filepath}")
    else:
        filepath = save_html_fallback(html_content, report_date)
        logger.info(f"フォールバック: {filepath}")

    logger.info("週次レポート処理が完了しました。")


if __name__ == "__main__":
    main()
