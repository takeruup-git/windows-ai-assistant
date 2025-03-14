import customtkinter as ctk
from openai import OpenAI
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import ollama
import os
import pickle
import threading
import requests
import json
import base64
import time
from datetime import datetime
import tkinter as tk
from tkinter import messagebox
import re

# スコープ設定
SCOPES = [
    'https://www.googleapis.com/auth/drive.readonly',
    'https://www.googleapis.com/auth/gmail.modify',
    'https://www.googleapis.com/auth/tasks'
]

# クライアント設定
client = OpenAI(api_key=os.getenv('OPENAI_API_KEY'))

# Google APIクライアント取得
def get_google_service(service_name, version):
    """Google APIサービスのクライアントを取得する"""
    creds = None
    # token.pickleからの認証情報の読み込み
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    
    # 認証情報が無効か存在しない場合は更新または作成
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        
        # 次回のために認証情報を保存
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    
    # サービスクライアントを構築して返す
    return build(service_name, version, credentials=creds)

# オフラインチェック
def is_online():
    """インターネット接続を確認する"""
    try:
        requests.get("https://www.google.com", timeout=5)
        return True
    except requests.ConnectionError:
        return False

# OpenAI APIまたはローカルLLMを使用して応答を生成
def get_ai_response(prompt, system_message="あなたは役立つAIアシスタントです。"):
    """AIからの応答を取得する（オンラインならOpenAI、オフラインならGemma）"""
    if is_online():
        try:
            response = client.chat.completions.create(
                model="gpt-4o",
                messages=[
                    {"role": "system", "content": system_message},
                    {"role": "user", "content": prompt}
                ]
            )
            return response.choices[0].message.content
        except Exception as e:
            print(f"OpenAI API エラー: {e}")
            # OpenAI APIでエラーが発生した場合はGemmaにフォールバック
            return get_local_llm_response(prompt, system_message)
    else:
        return get_local_llm_response(prompt, system_message)

# ローカルLLM（Gemma）からの応答を取得
def get_local_llm_response(prompt, system_message="あなたは役立つAIアシスタントです。"):
    """ローカルLLM（Gemma）からの応答を取得する"""
    try:
        response = ollama.chat(
            model='gemma:1b',
            messages=[
                {"role": "system", "content": system_message},
                {"role": "user", "content": prompt}
            ]
        )
        return response['message']['content']
    except Exception as e:
        print(f"ローカルLLMエラー: {e}")
        return "申し訳ありません。AIサービスに接続できませんでした。インターネット接続とLLMの状態を確認してください。"

# ドライブ検索と提案
def drive_search_and_suggest(query, app):
    """Googleドライブを検索し、結果に基づいて提案を生成する"""
    app.update_progress(10, "Googleドライブに接続中...")
    
    try:
        # Googleドライブサービスの取得
        drive_service = get_google_service('drive', 'v3')
        
        app.update_progress(30, "ファイルを検索中...")
        
        # ドライブ内のファイルを検索
        results = drive_service.files().list(
            q=f"fullText contains '{query}' and trashed=false",
            spaces='drive',
            fields="files(id, name, mimeType, webViewLink, description, createdTime, modifiedTime)",
            pageSize=10
        ).execute()
        
        items = results.get('files', [])
        
        if not items:
            app.update_progress(100, "完了")
            return "検索結果が見つかりませんでした。別のキーワードで試してください。"
        
        app.update_progress(60, "検索結果を分析中...")
        
        # 検索結果の整形
        files_info = []
        for item in items:
            file_info = {
                "名前": item['name'],
                "タイプ": item['mimeType'],
                "リンク": item.get('webViewLink', 'リンクなし'),
                "作成日": item.get('createdTime', '不明'),
                "更新日": item.get('modifiedTime', '不明')
            }
            files_info.append(file_info)
        
        # AIに提案を生成させる
        app.update_progress(80, "AIによる提案を生成中...")
        
        prompt = f"""
        以下はGoogleドライブの検索結果です。キーワード「{query}」に関連するファイルです：
        
        {json.dumps(files_info, ensure_ascii=False, indent=2)}
        
        これらのファイルについて以下の情報を提供してください：
        1. 最も関連性が高そうなファイル3つとその理由
        2. これらのファイルを使って何ができるか、具体的な提案
        3. 検索結果から見つからなかった可能性のある関連情報
        
        簡潔かつ具体的に回答してください。
        """
        
        suggestion = get_ai_response(prompt)
        
        app.update_progress(100, "完了")
        
        # 結果を返す
        result = f"### 検索結果: {len(items)}件のファイルが見つかりました\n\n"
        result += f"### AIによる提案:\n{suggestion}"
        
        return result
    
    except Exception as e:
        app.update_progress(100, "エラーが発生しました")
        return f"エラーが発生しました: {str(e)}"

# メール処理
def process_email(app):
    """未読メールを取得し、AIによる返信提案を生成する"""
    app.update_progress(10, "Gmailに接続中...")
    
    try:
        # Gmailサービスの取得
        gmail_service = get_google_service('gmail', 'v1')
        
        app.update_progress(30, "未読メールを取得中...")
        
        # 未読メールの取得
        results = gmail_service.users().messages().list(
            userId='me',
            q='is:unread',
            maxResults=5
        ).execute()
        
        messages = results.get('messages', [])
        
        if not messages:
            app.update_progress(100, "完了")
            return "未読メールはありません。"
        
        app.update_progress(50, "メール内容を分析中...")
        
        email_contents = []
        
        for message in messages:
            msg = gmail_service.users().messages().get(
                userId='me',
                id=message['id'],
                format='full'
            ).execute()
            
            # メールの件名を取得
            subject = ""
            sender = ""
            for header in msg['payload']['headers']:
                if header['name'] == 'Subject':
                    subject = header['value']
                if header['name'] == 'From':
                    sender = header['value']
            
            # メール本文を取得
            body = ""
            if 'parts' in msg['payload']:
                for part in msg['payload']['parts']:
                    if part['mimeType'] == 'text/plain':
                        body = base64.urlsafe_b64decode(part['body']['data']).decode('utf-8')
                        break
            elif 'body' in msg['payload'] and 'data' in msg['payload']['body']:
                body = base64.urlsafe_b64decode(msg['payload']['body']['data']).decode('utf-8')
            
            email_contents.append({
                "id": message['id'],
                "subject": subject,
                "sender": sender,
                "body": body[:500] + ("..." if len(body) > 500 else "")  # 長すぎる場合は省略
            })
        
        app.update_progress(70, "AIによる返信提案を生成中...")
        
        # 各メールに対する返信提案を生成
        email_responses = []
        
        for email in email_contents:
            prompt = f"""
            以下のメールに対する適切な返信を日本語で提案してください：
            
            差出人: {email['sender']}
            件名: {email['subject']}
            本文:
            {email['body']}
            
            返信内容は簡潔かつ丁寧に、ビジネスメールとして適切な形式で作成してください。
            """
            
            response = get_ai_response(prompt)
            
            email_responses.append({
                "id": email['id'],
                "subject": email['subject'],
                "sender": email['sender'],
                "response": response
            })
        
        app.update_progress(100, "完了")
        
        # 結果を整形
        result = f"### 未読メール: {len(messages)}件\n\n"
        
        for i, resp in enumerate(email_responses):
            result += f"## メール {i+1}\n"
            result += f"差出人: {resp['sender']}\n"
            result += f"件名: {resp['subject']}\n\n"
            result += f"### 提案される返信:\n{resp['response']}\n\n"
            result += "---\n\n"
        
        return result
    
    except Exception as e:
        app.update_progress(100, "エラーが発生しました")
        return f"エラーが発生しました: {str(e)}"

# タスク追加
def add_task_from_content(content, app):
    """テキスト内容からタスクを抽出し、Googleタスクに追加する"""
    app.update_progress(10, "内容を分析中...")
    
    try:
        # AIを使用してタスクを抽出
        prompt = f"""
        以下のテキストからタスクを抽出してください。各タスクには以下の情報を含めてください：
        1. タスクのタイトル（簡潔に）
        2. タスクの詳細説明
        3. 推定される締め切り（テキストから推測できる場合）
        4. 優先度（高/中/低）
        
        JSON形式で返してください。例：
        [
          {{
            "title": "会議の準備",
            "notes": "プレゼン資料を作成し、参加者に送付する",
            "due": "2023-12-15",
            "priority": "高"
          }}
        ]
        
        テキスト:
        {content}
        """
        
        app.update_progress(30, "AIによるタスク抽出中...")
        
        tasks_json_text = get_ai_response(prompt)
        
        # JSON部分を抽出（AIの回答にテキストが含まれている場合に対応）
        json_match = re.search(r'\[\s*\{.*\}\s*\]', tasks_json_text, re.DOTALL)
        if json_match:
            tasks_json_text = json_match.group(0)
        
        try:
            tasks = json.loads(tasks_json_text)
        except json.JSONDecodeError:
            app.update_progress(100, "エラーが発生しました")
            return "タスクの抽出に失敗しました。テキスト形式を確認してください。"
        
        if not tasks:
            app.update_progress(100, "完了")
            return "タスクが見つかりませんでした。別のテキストで試してください。"
        
        app.update_progress(60, "Googleタスクに接続中...")
        
        # Googleタスクサービスの取得
        tasks_service = get_google_service('tasks', 'v1')
        
        # タスクリストの取得
        task_lists = tasks_service.tasklists().list().execute()
        
        if not task_lists or 'items' not in task_lists:
            # タスクリストがない場合は新規作成
            tasklist = tasks_service.tasklists().insert(body={
                'title': 'AIアシスタント'
            }).execute()
            tasklist_id = tasklist['id']
        else:
            # 既存のタスクリストを使用（最初のもの）
            tasklist_id = task_lists['items'][0]['id']
        
        app.update_progress(80, "タスクを追加中...")
        
        # タスクの追加
        added_tasks = []
        
        for task in tasks:
            # 締め切りの形式を調整
            due_date = None
            if 'due' in task and task['due']:
                try:
                    # YYYY-MM-DD形式に変換
                    parsed_date = datetime.strptime(task['due'], '%Y-%m-%d')
                    due_date = parsed_date.strftime('%Y-%m-%dT00:00:00.000Z')
                except ValueError:
                    # 日付形式が異なる場合は無視
                    pass
            
            # タスクの作成
            task_body = {
                'title': task['title'],
                'notes': task.get('notes', '')
            }
            
            if due_date:
                task_body['due'] = due_date
            
            result = tasks_service.tasks().insert(
                tasklist=tasklist_id,
                body=task_body
            ).execute()
            
            added_tasks.append({
                'title': task['title'],
                'notes': task.get('notes', ''),
                'due': task.get('due', '未設定'),
                'priority': task.get('priority', '中'),
                'id': result['id']
            })
        
        app.update_progress(100, "完了")
        
        # 結果を整形
        result = f"### 追加されたタスク: {len(added_tasks)}件\n\n"
        
        for i, task in enumerate(added_tasks):
            result += f"## タスク {i+1}: {task['title']}\n"
            result += f"詳細: {task['notes']}\n"
            result += f"期限: {task['due']}\n"
            result += f"優先度: {task['priority']}\n\n"
        
        return result
    
    except Exception as e:
        app.update_progress(100, "エラーが発生しました")
        return f"エラーが発生しました: {str(e)}"

# Webレポート生成
def generate_web_report(topic, app):
    """指定されたトピックに関するWebレポートを生成する"""
    app.update_progress(10, "情報収集を開始...")
    
    try:
        if not is_online():
            app.update_progress(20, "オフラインモード: ローカルデータのみ使用...")
            
            # ドライブからの情報収集
            drive_service = get_google_service('drive', 'v3')
            
            results = drive_service.files().list(
                q=f"fullText contains '{topic}' and trashed=false",
                spaces='drive',
                fields="files(id, name, mimeType, description)",
                pageSize=5
            ).execute()
            
            items = results.get('files', [])
            
            files_info = []
            for item in items:
                file_info = {
                    "名前": item['name'],
                    "タイプ": item['mimeType'],
                    "説明": item.get('description', '説明なし')
                }
                files_info.append(file_info)
            
            prompt = f"""
            以下は「{topic}」に関するGoogleドライブ内のファイル情報です：
            
            {json.dumps(files_info, ensure_ascii=False, indent=2)}
            
            これらの情報を元に、「{topic}」に関する5ページ程度のレポートを作成してください。
            レポートには以下のセクションを含めてください：
            
            1. 概要
            2. 背景
            3. 主要なポイント（3-5つ）
            4. 分析と考察
            5. 結論と提言
            
            注：このレポートはオフラインモードで生成されており、最新のWeb情報は含まれていません。
            """
            
            app.update_progress(50, "レポートを生成中...")
            
            report = get_local_llm_response(prompt)
            
        else:
            app.update_progress(20, "Web情報を検索中...")
            
            # OpenAIのWeb検索機能を使用
            web_search_prompt = f"""
            「{topic}」について詳細な情報を収集してください。以下の点に注目してください：
            1. 最新の動向や統計
            2. 主要な課題や機会
            3. 業界の専門家の見解
            4. 将来の展望
            
            これらの情報を元に、包括的なレポートを作成するための情報を集めています。
            """
            
            try:
                web_search_response = client.chat.completions.create(
                    model="gpt-4o",
                    messages=[
                        {"role": "system", "content": "あなたはWeb検索機能を持つAIアシスタントです。最新の情報を収集して提供してください。"},
                        {"role": "user", "content": web_search_prompt}
                    ]
                )
                web_info = web_search_response.choices[0].message.content
            except Exception as e:
                print(f"Web検索エラー: {e}")
                web_info = "Web検索に失敗しました。ローカルデータのみを使用します。"
            
            app.update_progress(40, "ドライブ情報を収集中...")
            
            # ドライブからの情報収集
            drive_service = get_google_service('drive', 'v3')
            
            results = drive_service.files().list(
                q=f"fullText contains '{topic}' and trashed=false",
                spaces='drive',
                fields="files(id, name, mimeType, description)",
                pageSize=5
            ).execute()
            
            items = results.get('files', [])
            
            files_info = []
            for item in items:
                file_info = {
                    "名前": item['name'],
                    "タイプ": item['mimeType'],
                    "説明": item.get('description', '説明なし')
                }
                files_info.append(file_info)
            
            app.update_progress(60, "レポートを生成中...")
            
            # レポート生成
            report_prompt = f"""
            以下は「{topic}」に関する情報です：
            
            ## Web検索結果:
            {web_info}
            
            ## Googleドライブ内の関連ファイル:
            {json.dumps(files_info, ensure_ascii=False, indent=2)}
            
            これらの情報を元に、「{topic}」に関する5ページ程度の包括的なレポートを作成してください。
            レポートには以下のセクションを含めてください：
            
            1. エグゼクティブサマリー
            2. 背景と市場概況
            3. 主要な発見（4-6つ）
            4. 詳細分析
            5. 結論と戦略的提言
            6. 参考資料
            
            各セクションは見出しを付け、内容は具体的かつ実用的にしてください。
            """
            
            report = get_ai_response(report_prompt)
        
        app.update_progress(100, "完了")
        
        return f"### 「{topic}」に関するレポート\n\n{report}"
    
    except Exception as e:
        app.update_progress(100, "エラーが発生しました")
        return f"エラーが発生しました: {str(e)}"

# GUI設定
class AIAssistantApp:
    def __init__(self, root):
        self.root = root
        self.root.title("AI業務アシスタント")
        self.root.geometry("800x600")
        
        # テーマ設定
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        
        # メインフレーム
        self.main_frame = ctk.CTkFrame(self.root)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # タイトル
        self.title_label = ctk.CTkLabel(
            self.main_frame, 
            text="AI業務アシスタント", 
            font=ctk.CTkFont(size=20, weight="bold")
        )
        self.title_label.pack(pady=10)
        
        # 入力フレーム
        self.input_frame = ctk.CTkFrame(self.main_frame)
        self.input_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.input_label = ctk.CTkLabel(self.input_frame, text="クエリ/内容:")
        self.input_label.pack(side=tk.LEFT, padx=5)
        
        self.input_text = ctk.CTkEntry(self.input_frame, width=500)
        self.input_text.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # ボタンフレーム
        self.button_frame = ctk.CTkFrame(self.main_frame)
        self.button_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.drive_button = ctk.CTkButton(
            self.button_frame, 
            text="ドライブ検索", 
            command=self.on_drive_search
        )
        self.drive_button.pack(side=tk.LEFT, padx=5, expand=True)
        
        self.email_button = ctk.CTkButton(
            self.button_frame, 
            text="メール処理", 
            command=self.on_email_process
        )
        self.email_button.pack(side=tk.LEFT, padx=5, expand=True)
        
        self.task_button = ctk.CTkButton(
            self.button_frame, 
            text="タスク追加", 
            command=self.on_task_add
        )
        self.task_button.pack(side=tk.LEFT, padx=5, expand=True)
        
        self.report_button = ctk.CTkButton(
            self.button_frame, 
            text="レポート生成", 
            command=self.on_report_generate
        )
        self.report_button.pack(side=tk.LEFT, padx=5, expand=True)
        
        # 進捗バー
        self.progress_frame = ctk.CTkFrame(self.main_frame)
        self.progress_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.progress_bar = ctk.CTkProgressBar(self.progress_frame)
        self.progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.progress_bar.set(0)
        
        self.status_label = ctk.CTkLabel(self.progress_frame, text="待機中...")
        self.status_label.pack(side=tk.LEFT, padx=5)
        
        # 結果表示エリア
        self.result_frame = ctk.CTkFrame(self.main_frame)
        self.result_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.result_label = ctk.CTkLabel(self.result_frame, text="結果:")
        self.result_label.pack(anchor=tk.W, padx=5, pady=5)
        
        self.result_text = ctk.CTkTextbox(self.result_frame, wrap=tk.WORD)
        self.result_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 初期状態設定
        self.update_progress(0, "待機中...")
    
    def update_progress(self, value, status_text):
        """進捗バーと状態テキストを更新する"""
        self.progress_bar.set(value / 100)
        self.status_label.configure(text=status_text)
        self.root.update_idletasks()
    
    def set_result(self, text):
        """結果テキストを設定する"""
        self.result_text.delete("0.0", tk.END)
        self.result_text.insert("0.0", text)
    
    def on_drive_search(self):
        """ドライブ検索ボタンのイベントハンドラ"""
        query = self.input_text.get().strip()
        if not query:
            messagebox.showwarning("入力エラー", "検索クエリを入力してください。")
            return
        
        # ボタンを無効化
        self.disable_buttons()
        
        # 別スレッドで処理を実行
        threading.Thread(target=self.run_drive_search, args=(query,), daemon=True).start()
    
    def run_drive_search(self, query):
        """ドライブ検索を実行する"""
        result = drive_search_and_suggest(query, self)
        self.set_result(result)
        self.enable_buttons()
    
    def on_email_process(self):
        """メール処理ボタンのイベントハンドラ"""
        # ボタンを無効化
        self.disable_buttons()
        
        # 別スレッドで処理を実行
        threading.Thread(target=self.run_email_process, daemon=True).start()
    
    def run_email_process(self):
        """メール処理を実行する"""
        result = process_email(self)
        self.set_result(result)
        self.enable_buttons()
    
    def on_task_add(self):
        """タスク追加ボタンのイベントハンドラ"""
        content = self.input_text.get().strip()
        if not content:
            messagebox.showwarning("入力エラー", "タスクを抽出するテキストを入力してください。")
            return
        
        # ボタンを無効化
        self.disable_buttons()
        
        # 別スレッドで処理を実行
        threading.Thread(target=self.run_task_add, args=(content,), daemon=True).start()
    
    def run_task_add(self, content):
        """タスク追加を実行する"""
        result = add_task_from_content(content, self)
        self.set_result(result)
        self.enable_buttons()
    
    def on_report_generate(self):
        """レポート生成ボタンのイベントハンドラ"""
        topic = self.input_text.get().strip()
        if not topic:
            messagebox.showwarning("入力エラー", "レポートのトピックを入力してください。")
            return
        
        # ボタンを無効化
        self.disable_buttons()
        
        # 別スレッドで処理を実行
        threading.Thread(target=self.run_report_generate, args=(topic,), daemon=True).start()
    
    def run_report_generate(self, topic):
        """レポート生成を実行する"""
        result = generate_web_report(topic, self)
        self.set_result(result)
        self.enable_buttons()
    
    def disable_buttons(self):
        """ボタンを無効化する"""
        self.drive_button.configure(state=tk.DISABLED)
        self.email_button.configure(state=tk.DISABLED)
        self.task_button.configure(state=tk.DISABLED)
        self.report_button.configure(state=tk.DISABLED)
    
    def enable_buttons(self):
        """ボタンを有効化する"""
        self.drive_button.configure(state=tk.NORMAL)
        self.email_button.configure(state=tk.NORMAL)
        self.task_button.configure(state=tk.NORMAL)
        self.report_button.configure(state=tk.NORMAL)

# メイン実行
def setup_gui():
    """GUIをセットアップして実行する"""
    root = ctk.CTk()
    app = AIAssistantApp(root)
    root.mainloop()

if __name__ == "__main__":
    setup_gui()