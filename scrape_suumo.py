#!/usr/bin/env python3
"""
SUUMOスクレイピングスクリプト
上本町ハイハイタウン周辺の中古マンション情報を取得する
"""

import json
import logging
import re
import time
import urllib.parse
from datetime import datetime
from pathlib import Path
from typing import Any

import requests
from bs4 import BeautifulSoup

# ログ設定
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(),
    ],
)
logger = logging.getLogger(__name__)

# 定数
SCRIPT_DIR = Path(__file__).parent
TODAY = datetime.now().strftime("%Y-%m-%d")
TODAY_COMPACT = datetime.now().strftime("%Y%m%d")
OUTPUT_FILE = SCRIPT_DIR / f"suumo_data_{TODAY_COMPACT}.json"

# リクエスト設定
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}

# リクエスト間の待機秒数（サーバーへの負荷軽減）
REQUEST_DELAY_SEC = 3.0
# リトライ回数
MAX_RETRIES = 3
# リトライ間隔（秒）
RETRY_DELAY_SEC = 5.0

# SUUMO検索URL
# ハイハイタウン名前検索
HAIHAITOWN_URL = (
    "https://suumo.jp/jj/bukken/ichiran/JJ010FJ001/"
    "?ar=060&bs=011&fw=%E3%83%8F%E3%82%A4%E3%83%8F%E3%82%A4%E3%82%BF%E3%82%A6%E3%83%B3"
    "&po=0&pj=0"
)

# 天王寺区の中古マンション一覧
TENNOJI_AREA_URL = (
    "https://suumo.jp/jj/bukken/ichiran/JJ010FJ001/"
    "?ar=060&bs=011&ta=27&sc=27109"
    "&kb=1&kt=9999&mb=0&mt=9999999"
    "&ekTjCd=&ekTjNm=&tj=0&cnb=0&cn=9999999"
)


def fetch_page(url: str, session: requests.Session) -> BeautifulSoup | None:
    """
    指定URLのHTMLを取得してBeautifulSoupオブジェクトを返す。
    リトライ付き。
    """
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            logger.info("ページ取得中 (試行 %d/%d): %s", attempt, MAX_RETRIES, url[:120])
            resp = session.get(url, headers=HEADERS, timeout=30)
            resp.raise_for_status()
            resp.encoding = resp.apparent_encoding or "utf-8"
            return BeautifulSoup(resp.text, "html.parser")
        except requests.RequestException as e:
            logger.warning("リクエスト失敗 (試行 %d/%d): %s", attempt, MAX_RETRIES, e)
            if attempt < MAX_RETRIES:
                logger.info("%.1f秒後にリトライします...", RETRY_DELAY_SEC)
                time.sleep(RETRY_DELAY_SEC)
            else:
                logger.error("最大リトライ回数に到達。スキップします: %s", url[:120])
                return None


def extract_number(text: str) -> float | None:
    """テキストから数値を抽出する"""
    if not text:
        return None
    # 全角数字を半角に変換
    text = text.translate(str.maketrans("０１２３４５６７８９．", "0123456789."))
    match = re.search(r"[\d,.]+", text.replace(",", ""))
    if match:
        try:
            return float(match.group())
        except ValueError:
            return None
    return None


def parse_listing_from_cassetteitem(item) -> dict[str, Any] | None:
    """
    SUUMOの物件カセット（.property_unit / .cassetteitem）から
    物件情報を抽出する。
    SUUMOのHTML構造はページによって異なるため、複数パターンに対応。
    """
    listing: dict[str, Any] = {}

    try:
        # --- 物件名 ---
        name_el = (
            item.select_one(".property_unit-title a")
            or item.select_one(".cassetteitem_content-title")
            or item.select_one("a[href*='/ms/']")  # マンション詳細リンク
        )
        if name_el:
            listing["name"] = name_el.get_text(strip=True)
            # 詳細URL
            href = name_el.get("href", "")
            if href and not href.startswith("http"):
                href = "https://suumo.jp" + href
            listing["url"] = href
        else:
            # 物件名が取れなければスキップ
            return None

        # --- テーブル形式（property_unit-body内のdl/dt/dd） ---
        dts = item.select("dt")
        dds = item.select("dd")
        info_map: dict[str, str] = {}
        for dt_el, dd_el in zip(dts, dds):
            key = dt_el.get_text(strip=True)
            val = dd_el.get_text(strip=True)
            info_map[key] = val

        # --- 価格 ---
        price_text = info_map.get("販売価格", "") or info_map.get("価格", "")
        if not price_text:
            price_el = item.select_one(".dottable-value--price") or item.select_one(
                ".cassetteitem_price--accent"
            )
            if price_el:
                price_text = price_el.get_text(strip=True)
        price_raw = extract_number(price_text)
        # 「万円」表記の場合
        if price_raw is not None:
            if "億" in price_text:
                # 「1億2000万円」形式を処理
                oku_match = re.search(r"([\d.]+)\s*億", price_text)
                man_match = re.search(r"億\s*([\d,.]+)\s*万", price_text)
                price_man = 0.0
                if oku_match:
                    price_man += float(oku_match.group(1)) * 10000
                if man_match:
                    price_man += float(man_match.group(1).replace(",", ""))
                listing["price_man"] = price_man
            else:
                listing["price_man"] = price_raw
        else:
            listing["price_man"] = None

        # --- 専有面積 ---
        area_text = info_map.get("専有面積", "")
        if not area_text:
            area_el = item.select_one(".dottable-value--area") or item.find(
                string=re.compile(r"[\d.]+m")
            )
            if area_el:
                area_text = (
                    area_el.get_text(strip=True)
                    if hasattr(area_el, "get_text")
                    else str(area_el)
                )
        area_sqm = extract_number(area_text)
        listing["area_sqm"] = area_sqm

        # --- ㎡単価を計算 ---
        if listing["price_man"] and area_sqm and area_sqm > 0:
            listing["price_per_sqm"] = round(listing["price_man"] / area_sqm, 2)
        else:
            listing["price_per_sqm"] = None

        # --- 間取り ---
        layout = info_map.get("間取り", "")
        if not layout:
            layout_el = item.select_one(".dottable-value--layout") or item.find(
                string=re.compile(r"\d[LDKS]+")
            )
            if layout_el:
                layout = (
                    layout_el.get_text(strip=True)
                    if hasattr(layout_el, "get_text")
                    else str(layout_el).strip()
                )
        listing["layout"] = layout or None

        # --- 階数 ---
        floor_text = info_map.get("所在階", "") or info_map.get("階", "")
        if not floor_text:
            floor_el = item.find(string=re.compile(r"\d+階"))
            if floor_el:
                floor_text = str(floor_el).strip()
        listing["floor"] = floor_text or None

        # --- 築年数 ---
        age_text = info_map.get("築年月", "") or info_map.get("築年数", "")
        if not age_text:
            age_el = item.find(string=re.compile(r"築\d+年|19\d{2}年|20\d{2}年"))
            if age_el:
                age_text = str(age_el).strip()
        listing["building_age"] = age_text or None

        # --- 最寄り駅 ---
        station_text = info_map.get("沿線・駅", "") or info_map.get("アクセス", "")
        if not station_text:
            station_el = item.select_one(
                ".cassetteitem_detail-col2"
            ) or item.select_one(".dottable-value--station")
            if station_el:
                station_text = station_el.get_text(strip=True)
        listing["station_access"] = station_text or None

        return listing

    except Exception as e:
        logger.debug("物件パース中にエラー: %s", e)
        return None


def parse_listing_page(soup: BeautifulSoup) -> list[dict[str, Any]]:
    """
    検索結果ページから全物件を抽出する。
    SUUMOの一覧ページは複数のHTML構造パターンがあるため、
    複数のセレクタで試行する。
    """
    listings: list[dict[str, Any]] = []

    # パターン1: property_unit（中古マンション一覧でよく使われる）
    items = soup.select(".property_unit")

    # パターン2: cassetteitem（賃貸などでよく使われるが中古にも出る）
    if not items:
        items = soup.select(".cassetteitem")

    # パターン3: js-bukkenList配下
    if not items:
        items = soup.select(".js-bukkenList > div")

    # パターン4: dottable-body内の各行
    if not items:
        items = soup.select(".dottable-body")

    logger.info("ページ内で %d 件の物件ブロックを検出", len(items))

    for item in items:
        listing = parse_listing_from_cassetteitem(item)
        if listing:
            listings.append(listing)

    return listings


def get_next_page_url(soup: BeautifulSoup, current_url: str) -> str | None:
    """ページネーションから次のページURLを取得する"""
    # 「次へ」リンクを探す
    next_link = soup.select_one(".pagination-parts a[rel='next']")
    if not next_link:
        next_link = soup.select_one("a.pagination_set-arrow--next")
    if not next_link:
        # テキストで「次へ」を探す
        for a_tag in soup.select(".pagination a, .paginate a"):
            if "次へ" in a_tag.get_text():
                next_link = a_tag
                break

    if next_link:
        href = next_link.get("href", "")
        if href and not href.startswith("http"):
            href = "https://suumo.jp" + href
        return href
    return None


def scrape_listings(
    start_url: str,
    session: requests.Session,
    label: str,
    max_pages: int = 5,
) -> list[dict[str, Any]]:
    """
    指定URLから物件一覧をスクレイピングする。
    ページネーションにも対応（最大max_pagesページ）。
    """
    all_listings: list[dict[str, Any]] = []
    current_url: str | None = start_url

    for page_num in range(1, max_pages + 1):
        if not current_url:
            break

        logger.info("[%s] %dページ目を取得中...", label, page_num)
        soup = fetch_page(current_url, session)
        if soup is None:
            logger.warning("[%s] ページ取得失敗。中断します。", label)
            break

        page_listings = parse_listing_page(soup)
        logger.info("[%s] %dページ目: %d件取得", label, page_num, len(page_listings))
        all_listings.extend(page_listings)

        # 次のページがあるか確認
        if page_num < max_pages:
            current_url = get_next_page_url(soup, current_url)
            if current_url:
                logger.info("次ページ発見。%.1f秒待機後に取得...", REQUEST_DELAY_SEC)
                time.sleep(REQUEST_DELAY_SEC)
            else:
                logger.info("[%s] 最終ページに到達。", label)
        else:
            logger.info("[%s] 最大ページ数(%d)に到達。", label, max_pages)

    return all_listings


def calc_avg_price_per_sqm(listings: list[dict[str, Any]]) -> float | None:
    """物件リストの平均㎡単価を計算する"""
    values = [
        lst["price_per_sqm"]
        for lst in listings
        if lst.get("price_per_sqm") is not None
    ]
    if values:
        return round(sum(values) / len(values), 2)
    return None


def main() -> None:
    """メイン処理"""
    logger.info("=" * 60)
    logger.info("SUUMO スクレイピング開始")
    logger.info("対象: 上本町ハイハイタウン周辺 中古マンション")
    logger.info("日付: %s", TODAY)
    logger.info("=" * 60)

    session = requests.Session()

    # --- ハイハイタウンの物件を検索 ---
    logger.info("")
    logger.info("【1/2】ハイハイタウンの物件を検索中...")
    haihaitown_listings = scrape_listings(
        HAIHAITOWN_URL, session, label="ハイハイタウン", max_pages=3
    )
    logger.info("ハイハイタウン: 合計 %d 件取得", len(haihaitown_listings))

    # サーバーへの負荷軽減のため待機
    time.sleep(REQUEST_DELAY_SEC)

    # --- 天王寺区エリアの物件を検索 ---
    logger.info("")
    logger.info("【2/2】天王寺区エリアの中古マンションを検索中...")
    area_listings = scrape_listings(
        TENNOJI_AREA_URL, session, label="天王寺区", max_pages=5
    )
    logger.info("天王寺区エリア: 合計 %d 件取得", len(area_listings))

    # --- サマリー集計 ---
    summary = {
        "haihaitown_count": len(haihaitown_listings),
        "haihaitown_avg_price_per_sqm": calc_avg_price_per_sqm(haihaitown_listings),
        "area_count": len(area_listings),
        "area_avg_price_per_sqm": calc_avg_price_per_sqm(area_listings),
    }

    # --- 結果をJSON保存 ---
    result = {
        "scrape_date": TODAY,
        "haihaitown_listings": haihaitown_listings,
        "area_listings": area_listings,
        "summary": summary,
    }

    OUTPUT_FILE.write_text(
        json.dumps(result, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    logger.info("")
    logger.info("=" * 60)
    logger.info("スクレイピング完了")
    logger.info("出力ファイル: %s", OUTPUT_FILE)
    logger.info("--- サマリー ---")
    logger.info(
        "  ハイハイタウン: %d件 (平均㎡単価: %s万円)",
        summary["haihaitown_count"],
        summary["haihaitown_avg_price_per_sqm"] or "N/A",
    )
    logger.info(
        "  天王寺区エリア: %d件 (平均㎡単価: %s万円)",
        summary["area_count"],
        summary["area_avg_price_per_sqm"] or "N/A",
    )
    logger.info("=" * 60)


if __name__ == "__main__":
    main()
