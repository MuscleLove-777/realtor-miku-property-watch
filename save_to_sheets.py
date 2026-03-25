#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
save_to_sheets.py
最新のsuumo_data_*.jsonを読み込み、Googleスプレッドシートにデータを追記する。

認証: サービスアカウント（環境変数 GOOGLE_CREDENTIALS_JSON にJSON文字列を格納）
スプレッドシート名: 不動産ウォッチ_ハイハイタウン
"""

import os
import sys
import json
import glob
import logging
from datetime import datetime, date

import gspread
from google.oauth2.service_account import Credentials

# ───────────────────────────── ログ設定 ─────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
logger = logging.getLogger(__name__)

# ───────────────────────────── 定数 ─────────────────────────────
SPREADSHEET_NAME = "不動産ウォッチ_ハイハイタウン"

# 各シートの定義（シート名, ヘッダー行）
SHEET_DEFINITIONS = {
    "ハイハイタウン売出し": [
        "日付", "物件名", "価格(万円)", "面積(㎡)",
        "㎡単価(万円)", "階数", "間取り", "URL",
    ],
    "周辺相場": [
        "日付", "物件名", "最寄駅", "価格(万円)",
        "面積(㎡)", "㎡単価(万円)", "築年数", "URL",
    ],
    "価格推移サマリー": [
        "日付", "HT平均㎡単価", "周辺平均㎡単価",
        "HT売出件数", "周辺売出件数",
    ],
}

# Google API スコープ
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# ───────────────────────────── 認証 ─────────────────────────────
def get_gspread_client() -> gspread.Client:
    """サービスアカウント認証でgspreadクライアントを返す。"""
    creds_json_str = os.environ.get("GOOGLE_CREDENTIALS_JSON")
    if not creds_json_str:
        raise EnvironmentError(
            "環境変数 GOOGLE_CREDENTIALS_JSON が設定されていません。"
            "サービスアカウントのJSON鍵文字列を設定してください。"
        )

    creds_info = json.loads(creds_json_str)
    credentials = Credentials.from_service_account_info(creds_info, scopes=SCOPES)
    client = gspread.authorize(credentials)
    logger.info("Google認証に成功しました。")
    return client


# ───────────────────────────── スプレッドシート取得/作成 ─────────────────────────────
def get_or_create_spreadsheet(client: gspread.Client) -> gspread.Spreadsheet:
    """スプレッドシートを取得する。存在しなければ新規作成する。"""
    try:
        spreadsheet = client.open(SPREADSHEET_NAME)
        logger.info(f"既存スプレッドシートを取得: {SPREADSHEET_NAME}")
    except gspread.SpreadsheetNotFound:
        spreadsheet = client.create(SPREADSHEET_NAME)
        logger.info(f"スプレッドシートを新規作成: {SPREADSHEET_NAME}")

    return spreadsheet


def get_or_create_worksheet(
    spreadsheet: gspread.Spreadsheet,
    sheet_name: str,
    headers: list[str],
) -> gspread.Worksheet:
    """ワークシートを取得する。存在しなければ作成しヘッダーを設定する。"""
    try:
        worksheet = spreadsheet.worksheet(sheet_name)
        logger.info(f"既存ワークシート取得: {sheet_name}")
    except gspread.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(
            title=sheet_name, rows=1000, cols=len(headers)
        )
        logger.info(f"ワークシートを新規作成: {sheet_name}")

    # ヘッダーが空なら設定
    existing = worksheet.row_values(1)
    if not existing:
        worksheet.append_row(headers, value_input_option="USER_ENTERED")
        logger.info(f"ヘッダーを設定: {sheet_name}")

    return worksheet


# ───────────────────────────── JSONファイル読み込み ─────────────────────────────
def find_latest_json(directory: str) -> str:
    """指定ディレクトリから最新の suumo_data_*.json ファイルのパスを返す。"""
    pattern = os.path.join(directory, "suumo_data_*.json")
    files = glob.glob(pattern)
    if not files:
        raise FileNotFoundError(
            f"suumo_data_*.json が見つかりません: {directory}"
        )

    # ファイル名に含まれる日付 or 更新日時でソートして最新を取得
    latest = max(files, key=os.path.getmtime)
    logger.info(f"最新JSONファイル: {latest}")
    return latest


def load_json(filepath: str) -> dict:
    """JSONファイルを読み込んで辞書を返す。"""
    with open(filepath, "r", encoding="utf-8") as f:
        data = json.load(f)
    logger.info(f"JSONデータ読み込み完了: {len(data) if isinstance(data, list) else 'dict'}")
    return data


# ───────────────────────────── ㎡単価の計算 ─────────────────────────────
def calc_unit_price(price_man: float | None, area_sqm: float | None) -> float | None:
    """価格(万円)と面積(㎡)から㎡単価を計算する。"""
    if price_man and area_sqm and area_sqm > 0:
        return round(price_man / area_sqm, 2)
    return None


# ───────────────────────────── データ整形 ─────────────────────────────
def build_haihai_rows(data: dict, today_str: str) -> list[list]:
    """ハイハイタウン売出しデータの行を構築する。"""
    rows = []
    listings = data.get("haihai_listings", data.get("listings", []))
    if isinstance(data, list):
        listings = data

    for item in listings:
        price = item.get("price_man") or item.get("price")
        area = item.get("area_sqm") or item.get("area")
        unit_price = calc_unit_price(price, area)
        row = [
            today_str,
            item.get("name", item.get("building_name", "")),
            price,
            area,
            unit_price,
            item.get("floor", item.get("floors", "")),
            item.get("layout", item.get("madori", "")),
            item.get("url", item.get("detail_url", "")),
        ]
        rows.append(row)
    return rows


def build_nearby_rows(data: dict, today_str: str) -> list[list]:
    """周辺相場データの行を構築する。"""
    rows = []
    nearby = data.get("nearby_listings", data.get("nearby", []))

    for item in nearby:
        price = item.get("price_man") or item.get("price")
        area = item.get("area_sqm") or item.get("area")
        unit_price = calc_unit_price(price, area)
        row = [
            today_str,
            item.get("name", item.get("building_name", "")),
            item.get("station", item.get("nearest_station", "")),
            price,
            area,
            unit_price,
            item.get("age", item.get("building_age", "")),
            item.get("url", item.get("detail_url", "")),
        ]
        rows.append(row)
    return rows


def build_summary_row(
    data: dict, today_str: str, haihai_rows: list, nearby_rows: list
) -> list:
    """価格推移サマリーの行を構築する。"""
    # ハイハイタウンの平均㎡単価
    ht_prices = [r[4] for r in haihai_rows if r[4] is not None]
    ht_avg = round(sum(ht_prices) / len(ht_prices), 2) if ht_prices else None

    # 周辺の平均㎡単価
    nb_prices = [r[5] for r in nearby_rows if r[5] is not None]
    nb_avg = round(sum(nb_prices) / len(nb_prices), 2) if nb_prices else None

    return [
        today_str,
        ht_avg,
        nb_avg,
        len(haihai_rows),
        len(nearby_rows),
    ]


# ───────────────────────────── メイン処理 ─────────────────────────────
def main():
    """メイン処理: JSONを読みスプレッドシートに追記する。"""
    # スクリプトのあるディレクトリを基準にする
    script_dir = os.path.dirname(os.path.abspath(__file__))

    # 最新JSONファイルを探す
    try:
        json_path = find_latest_json(script_dir)
    except FileNotFoundError as e:
        logger.error(str(e))
        sys.exit(1)

    # データ読み込み
    data = load_json(json_path)
    today_str = date.today().isoformat()

    # Google認証
    try:
        client = get_gspread_client()
    except Exception as e:
        logger.error(f"Google認証エラー: {e}")
        sys.exit(1)

    # スプレッドシート取得/作成
    spreadsheet = get_or_create_spreadsheet(client)

    # デフォルトの「Sheet1」を削除（他シートが存在する場合のみ）
    try:
        default_sheet = spreadsheet.worksheet("Sheet1")
        if len(spreadsheet.worksheets()) > 1:
            spreadsheet.del_worksheet(default_sheet)
            logger.info("デフォルトの Sheet1 を削除しました。")
    except gspread.WorksheetNotFound:
        pass

    # ─── データ行の構築 ───
    haihai_rows = build_haihai_rows(data, today_str)
    nearby_rows = build_nearby_rows(data, today_str)
    summary_row = build_summary_row(data, today_str, haihai_rows, nearby_rows)

    # ─── 各シートにデータ追記 ───
    # (1) ハイハイタウン売出し
    ws_haihai = get_or_create_worksheet(
        spreadsheet,
        "ハイハイタウン売出し",
        SHEET_DEFINITIONS["ハイハイタウン売出し"],
    )
    if haihai_rows:
        ws_haihai.append_rows(haihai_rows, value_input_option="USER_ENTERED")
        logger.info(f"ハイハイタウン売出し: {len(haihai_rows)}行を追記")
    else:
        logger.warning("ハイハイタウン売出し: データなし")

    # (2) 周辺相場
    ws_nearby = get_or_create_worksheet(
        spreadsheet,
        "周辺相場",
        SHEET_DEFINITIONS["周辺相場"],
    )
    if nearby_rows:
        ws_nearby.append_rows(nearby_rows, value_input_option="USER_ENTERED")
        logger.info(f"周辺相場: {len(nearby_rows)}行を追記")
    else:
        logger.warning("周辺相場: データなし")

    # (3) 価格推移サマリー
    ws_summary = get_or_create_worksheet(
        spreadsheet,
        "価格推移サマリー",
        SHEET_DEFINITIONS["価格推移サマリー"],
    )
    ws_summary.append_row(summary_row, value_input_option="USER_ENTERED")
    logger.info("価格推移サマリー: 1行を追記")

    logger.info("すべてのデータ追記が完了しました。")
    logger.info(f"スプレッドシートURL: {spreadsheet.url}")


if __name__ == "__main__":
    main()
