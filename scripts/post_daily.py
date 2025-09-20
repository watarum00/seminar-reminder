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
    # スプレッドシートのメタデータ取得（全シートの title と index を取得）
    try:
        meta = service.spreadsheets().get(spreadsheetId=sheet_id, fields="sheets(properties(title,index))").execute()
    except Exception as e:
        raise RuntimeError(f"スプレッドシートのメタデータ取得に失敗: {e}")
    sheets = meta.get('sheets', [])
    if not sheets:
        raise RuntimeError("スプレッドシートにシートが存在しません")

    # 環境変数で参照したいシートを指定可能にする
    # 優先度: SHEET_NAME が最優先、次に SHEET_INDEX（0-based）。未指定なら先頭シートを使用
    sheet_name_env = os.environ.get('SHEET_NAME')
    sheet_index_env = os.environ.get('SHEET_INDEX')

    sheet_name = None
    if sheet_name_env:
        # 指定されたシート名をそのまま使用（存在確認は後続の API 呼び出しで失敗するためこの時点では軽く trim）
        sheet_name = sheet_name_env.strip()
    elif sheet_index_env:
        try:
            idx = int(sheet_index_env)
        except ValueError:
            raise RuntimeError(f"環境変数 SHEET_INDEX が整数ではありません: {sheet_index_env}")
        # sheets の中から index が一致するシートを探す
        matched = None
        for s in sheets:
            props = s.get('properties', {})
            if props.get('index') == idx:
                matched = props
                break
        if matched is None:
            raise RuntimeError(f"スプレッドシートに index={idx} のシートが見つかりません")
        sheet_name = matched.get('title')
    else:
        # デフォルト: 先頭シート
        first = sheets[0]['properties']
        sheet_name = first['title']
    # シート全体の値を取得（header も含む）
    try:
        # range にシート名を指定すると、そのシートの全セルを取得します
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
    # デバッグ出力: ヘッダと最初の数行
    if os.environ.get('DEBUG'):
        print(f"[DEBUG] Loaded {len(records)} records. Headers: {headers}")
        for i, r in enumerate(records[:10]):
            print(f"[DEBUG] record[{i}]: {r}")
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
    debug = bool(os.environ.get('DEBUG'))
    for row in records:
        if debug:
            print(f"[DEBUG] row raw: {row}")
        date_str = row.get('日付') or row.get('date') or row.get('Date') or ''
        time_str = row.get('テスト時間') or row.get('time') or row.get('Time') or ''
        # 予定のタイプ（優先: 予定タイプ > タイプ > 種別 > type）
        type_str = (row.get('予定タイプ') or row.get('タイプ') or row.get('種別') or row.get('type') or row.get('Type') or '').strip()
        content = row.get('内容') or row.get('content') or ''
        person = row.get('担当') or row.get('person') or ''

        # 欠席者列（優先: 欠席予定 > 欠席者 > 欠席 > absent）
        absent_raw = (row.get('欠席予定') or row.get('欠席者') or row.get('欠席') or row.get('absent') or row.get('Absentees') or '')
        absent_raw = str(absent_raw).strip()

        if not date_str:
            if debug:
                print(f"[DEBUG] skipping row because date_str is empty: {row}")
            continue
        d = parse_date_str(str(date_str).strip(), reference_year)
        if d is None:
            if debug:
                print(f"[DEBUG] could not parse date from '{date_str}' (ref year {reference_year})")
            continue

        if d in week_dates:
            if debug:
                print(f"[DEBUG] Found event on {d}: {content} (time: {time_str}, person: {person}, type: {type_str})")
            # type_str の扱い:
            # - 空文字: 何もしない
            # - 'ゼミ' : ゼミ扱い（time がある場合のみ追加）
            # - それ以外: 未知の値は '重要日程' 扱いにする（時間表示や注意は出さない）
            if type_str == 'ゼミ':
                ev_type = 'ゼミ'
            else:
                ev_type = '重要日程'

            # 欠席者表示: 要求により、'欠席予定' の文字列をそのまま出力します（分割しない）
            absent_display = absent_raw if absent_raw else None

            ev = {
                'date': d,
                'time': time_str if time_str else None,
                'content': content,
                'person': person,
                'row': row,
                'type': ev_type,
                'type_raw': type_str,
                'absent_raw': absent_raw,
                'absent_display': absent_display,
            }

            if ev['type'] == 'ゼミ':
                if ev['time']:
                    events.append(ev)
                    if debug:
                        print(f"[DEBUG] Added event (ゼミ): {events[-1]}")
                else:
                    if debug:
                        print(f"[DEBUG] Skip event (ゼミ) without time: {content}")
            else:
                # 重要日程扱い（未知のタイプもここに入る）
                events.append(ev)
                if debug:
                    print(f"[DEBUG] Added event (重要日程): {events[-1]}")
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
        # タイプ別の表示
        if ev.get('type') == 'ゼミ':
            # デフォルト（ゼミ）: 1行目にメイン情報（タイトル）、
            # その下に箇条書きで担当・時間・欠席予定・注意を表示
            header_line = f"{md}({wd}): {ev['content']}"
            lines.append(header_line)

            # 箇条書きで各フィールドを表示
            if ev.get('person'):
                lines.append(f"> • 担当: {ev['person']}")
            if ev.get('time'):
                lines.append(f"> • 時間: {ev['time']}")

            # 欠席予定はそのまま表示するが、箇条書きとして出す
            # 欠席がいない場合は 'なし' と表示する
            if ev.get('absent_display'):
                lines.append(f"> • 欠席予定: {ev['absent_display']}")
            else:
                lines.append("> • 欠席予定: なし")

            # 注意表示は時間が存在し、かつ通常時間と異なる場合に箇条書きで出す
            if ev.get('time') and ev.get('time') != '13:00-14:30':
                lines.append("> • *※注意! 時間が通常の 13:00-14:30 以外です*")
        else:
            # 重要日程は時間や注意の出力を行わず、内容のみを出力
            line = f"{md}({wd}): {ev['content']}"
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

    if os.environ.get('DEBUG'):
        print(f"[DEBUG] today: {today}, monday: {monday}, week_dates: {sorted(list(week_dates))}")

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
