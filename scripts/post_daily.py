#!/usr/bin/env python3
import os
import datetime
from dateutil import tz, parser as dateparser
import requests  
from slack_sdk import WebClient
from slack_sdk.errors import SlackApiError
from googleapiclient.discovery import build
from dotenv import load_dotenv
load_dotenv()

def get_today_jst():
    """Asia/Tokyo の現在日時を返す datetime.datetime (tz-aware)。"""
    jst = tz.gettz('Asia/Tokyo')
    return datetime.datetime.now(tz=jst)

def get_monday_date(today=None):
    if today is None:
        today = get_today_jst()
    delta_days = today.isoweekday() - 1
    monday = today - datetime.timedelta(days=delta_days)
    return monday.date()

def build_week_dates(monday_date):
    return { monday_date + datetime.timedelta(days=i) for i in range(7) }

def load_public_sheet_records():
    """
    環境変数 GOOGLE_API_KEY と SHEET_ID が設定されており、
    該当スプレッドシートが「公開（Anyone with link can view）」になっている前提で、
    Google Sheets API を API Key で呼び出し、最初のシートの全データを辞書リストとして返す。
    """
    api_key = os.environ.get('GOOGLE_API_KEY')
    sheet_id = os.environ.get('SHEET_ID')
    if not api_key or not sheet_id:
        raise RuntimeError("環境変数 GOOGLE_API_KEY または SHEET_ID が設定されていません")
    # Sheets API クライアントを API キー付きで構築
    service = build('sheets', 'v4', developerKey=api_key)
    # スプレッドシートのメタデータ取得→最初のシート名取得
    try:
        meta = service.spreadsheets().get(spreadsheetId=sheet_id, fields="sheets(properties(title,index))").execute()
    except Exception as e:
        raise RuntimeError(f"スプレッドシートのメタデータ取得に失敗: {e}")
    sheets = meta.get('sheets', [])
    if not sheets:
        raise RuntimeError("スプレッドシートにシートが存在しません")
    # ここでは index=0 のシートを使う
    first = sheets[0]['properties']
    sheet_name = first['title']
    # シート全体の値を取得（header も含む）
    try:
        result = service.spreadsheets().values().get(
            spreadsheetId=sheet_id,
            range=sheet_name
        ).execute()
    except Exception as e:
        raise RuntimeError(f"スプレッドシートのデータ取得に失敗: {e}")
    values = result.get('values', [])
    if not values:
        return []
    # 1行目をヘッダーとみなし、以降を辞書化
    headers = values[0]
    records = []
    for row in values[1:]:
        rec = {}
        for i, h in enumerate(headers):
            # 同じヘッダー名が重複していない前提
            rec[h] = row[i] if i < len(row) else ''
        records.append(rec)
        # print(f"Loaded record: {rec}")
    return records

def parse_date_str(date_str, reference_year):
    if not date_str or not isinstance(date_str, str):
        return None
    # MM/DD 形式の場合は正規表現で明示的に処理すると安全:
    import re
    m = re.match(r'^\s*(\d{1,2})[/-](\d{1,2})(?:[/-]\d{2,4})?', date_str)
    if m:
        # もし年が文字列に含まれていない可能性を厳密に扱いたい場合は
        # さらに年付き形式との区別ロジックを追加可能。ただここでは dateparser も利用。
        try:
            month = int(m.group(1))
            day = int(m.group(2))
            return datetime.date(reference_year, month, day)
        except Exception:
            pass
    # fallback: dateutil.parser.parse
    try:
        dt = dateparser.parse(date_str, default=datetime.datetime(reference_year, 1, 1))
        return dt.date()
    except (ValueError, OverflowError):
        return None

def find_week_events(records, week_dates):
    events = []
    monday = min(week_dates)
    reference_year = monday.year
    for row in records:
        date_str = row.get('日付') or row.get('date') or row.get('Date') or ''
        time_str = row.get('テスト時間') or row.get('time') or row.get('Time') or ''
        content = row.get('内容') or row.get('content') or ''
        person = row.get('担当') or row.get('person') or ''
        if not date_str:
            continue
        d = parse_date_str(str(date_str).strip(), reference_year)
        if d is None:
            continue
        if d in week_dates:
            print(f"Found event on {d}: {content} (time: {time_str}, person: {person})")
            if time_str:
                events.append({
                    'date': d,
                    'time': time_str,
                    'content': content,
                    'person': person,
                    'row': row,
                })
                print(f"Added event: {events[-1]}")
    return events

def format_schedule(events, monday_date, week_dates):
    weekday_map = {0:'月',1:'火',2:'水',3:'木',4:'金',5:'土',6:'日'}
    start = monday_date
    end = monday_date + datetime.timedelta(days=6)
    # 表示を日本語月日表記にする例
    header = f"今週の予定：{start.month}月{start.day}日 〜 {end.month}月{end.day}日"
    if not events:
        return f"{header}\n予定はありません。"
    events_sorted = sorted(events, key=lambda ev: (ev['date'], ev.get('time') or ''))
    lines = [header]
    for ev in events_sorted:
        d = ev['date']
        md = f"{d.month}/{d.day}"
        wd = weekday_map[d.weekday()]
        line = f"{md}({wd}): {ev['content']}"
        if ev.get('person'):
            line += f" - 担当: {ev['person']}"
        if ev.get('time'):
            line += f" - 時間: {ev['time']}"
        if ev.get('time') != '13:00-14:30':
            line += f"\n*※注意! 時間が通常の 13:00-14:30 以外です*"
        lines.append(line)
    return "\n".join(lines)

def post_to_slack(text):
    token = os.environ.get('SLACK_BOT_TOKEN')
    channel = os.environ.get('SLACK_CHANNEL')
    print("Slack へ投稿します, channel:", channel)
    if not token or not channel:
        print("Slack 設定がされていないため、出力結果:")
        print(text)
        return
    client = WebClient(token=token)
    try:
        client.chat_postMessage(channel=channel, text=text)
        print("Slack へ投稿しました。")
    except SlackApiError as e:
        print(f"Slack への投稿に失敗: {e.response['error']}")
        print("出力結果:")
        print(text)

def main():
    today = get_today_jst()
    monday = get_monday_date(today)
    week_dates = build_week_dates(monday)

    # 公開シート用読み込み
    try:
        records = load_public_sheet_records()
    except Exception as e:
        print(f"スプレッドシートの読み込みに失敗: {e}")
        return

    events = find_week_events(records, week_dates)
    text = format_schedule(events, monday, week_dates)
    print(text)  # デバッグ用にコンソール出力
    post_to_slack(text)

if __name__ == '__main__':
    main()
