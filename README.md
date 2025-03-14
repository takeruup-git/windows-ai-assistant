# Windows AI業務アシスタント

Windows 10 Pro向けのAI業務アシスタントアプリケーションです。GoogleドライブAPI、Gmail API、GoogleタスクAPI、およびOpenAI APIを使用して、日常の業務タスクを自動化・効率化します。

## 機能

1. **ドライブ検索と提案**: GoogleドライブAPIで検索し、OpenAIで関連ファイルの提案を生成します。
2. **メール処理**: Gmail APIで未読メールを取得し、AIによる返信提案を生成します。
3. **タスク追加**: テキスト内容からタスクを抽出し、Googleタスクに自動追加します。
4. **Webレポート**: OpenAI Web Search機能とドライブデータを組み合わせて、5ページ程度のレポートを生成します。

## セットアップ手順

### 前提条件

- Windows 10 Pro
- Python 3.8以上
- Google Cloud Platformアカウント
- OpenAIアカウント
- Ollama（ローカルLLM実行用）

### 1. リポジトリのクローン

```bash
git clone https://github.com/yourusername/windows-ai-assistant.git
cd windows-ai-assistant
```

### 2. 依存パッケージのインストール

```bash
pip install -r requirements.txt
```

### 3. Google API認証情報の設定

1. [Google Cloud Console](https://console.cloud.google.com/)にアクセスします。
2. 新しいプロジェクトを作成します。
3. APIライブラリから以下のAPIを有効化します：
   - Google Drive API
   - Gmail API
   - Tasks API
4. 認証情報ページで「認証情報を作成」→「OAuthクライアントID」を選択します。
5. アプリケーションの種類として「デスクトップアプリケーション」を選択します。
6. 作成された認証情報をダウンロードし、`credentials.json`としてプロジェクトのルートディレクトリに保存します。

### 4. OpenAI APIキーの設定

1. [OpenAIのダッシュボード](https://platform.openai.com/account/api-keys)からAPIキーを取得します。
2. 環境変数として設定します：

```bash
# Windows PowerShellの場合
$env:OPENAI_API_KEY="your-api-key-here"

# または、システム環境変数として永続的に設定することも可能です
```

### 5. Ollamaのセットアップ（オフライン機能用）

1. [Ollama公式サイト](https://ollama.ai/)からOllamaをダウンロードしてインストールします。
2. Gemma 1Bモデルをダウンロードします：

```bash
ollama pull gemma:1b
```

### 6. アプリケーションの実行

```bash
python ai_agent.py
```

初回実行時には、Googleアカウントへのアクセス許可を求められます。ブラウザが開き、認証フローが完了すると、`token.pickle`ファイルが生成され、以降の認証に使用されます。

## 使用方法

1. **ドライブ検索**:
   - 検索クエリを入力欄に入力し、「ドライブ検索」ボタンをクリックします。
   - 関連ファイルとAIによる提案が表示されます。

2. **メール処理**:
   - 「メール処理」ボタンをクリックします。
   - 未読メールとそれぞれに対するAIの返信提案が表示されます。

3. **タスク追加**:
   - タスクを含むテキストを入力欄に入力し、「タスク追加」ボタンをクリックします。
   - AIがテキストからタスクを抽出し、Googleタスクに追加します。

4. **レポート生成**:
   - レポートのトピックを入力欄に入力し、「レポート生成」ボタンをクリックします。
   - AIがWeb検索とドライブデータを組み合わせて、包括的なレポートを生成します。

## オフライン機能

インターネット接続がない場合、アプリケーションは自動的にオフラインモードに切り替わります：

- OpenAI APIの代わりにローカルのGemma 1Bモデルを使用します。
- Googleドライブの同期済みデータのみを使用します。
- Web検索機能は利用できませんが、ローカルデータに基づいたレポート生成は可能です。

## トラブルシューティング

- **認証エラー**: `token.pickle`ファイルを削除して、再度認証フローを実行してください。
- **APIエラー**: Google Cloud ConsoleでAPIが有効化されていることを確認してください。
- **Ollamaエラー**: Ollamaサービスが実行中であることを確認してください。

## 依存パッケージ

- openai
- google-auth-oauthlib
- google-auth-httplib2
- google-api-python-client
- customtkinter
- requests
- ollama

## ライセンス

MIT