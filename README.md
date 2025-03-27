# Google Workspace MCP Server

このプロジェクトは以下のリポジトリをベースに作成しています：
[epaproditus/google-workspace-mcp-server](https://github.com/epaproditus/google-workspace-mcp-server)

Google Workspace の機能（カレンダー、メール）を MCP サーバーとして提供するアプリケーション

## 機能

### Gmail 機能

- `list_emails`: 受信トレイから最近のメールをフィルタリング付きで一覧表示
- `search_emails`: Gmail クエリ構文を使用した高度なメール検索
- `send_email`: CC、BCC 対応のメール送信
- `modify_email`: メールラベルの管理（アーカイブ、ゴミ箱、既読/未読）

### カレンダー機能

- `list_events`: 日付範囲指定付きの予定一覧表示
- `create_event`: 参加者付きの予定作成
- `update_event`: 既存予定の更新
- `delete_event`: 予定の削除

## 必要条件

1. **Node.js**: Node.js 20 以上をインストール
2. **Google Cloud Console 設定**:
   - [Google Cloud Console](https://console.cloud.google.com/)にアクセス
   - 新規プロジェクトの作成または既存プロジェクトの選択
   - Gmail API と Google Calendar API の有効化:
     1. "APIs & Services" > "Library"に移動
     2. "Gmail API"を検索して有効化
     3. "Google Calendar API"を検索して有効化
   - OAuth 2.0 認証情報の設定:
     1. "APIs & Services" > "Credentials"に移動
     2. "Create Credentials" > "OAuth client ID"をクリック
     3. "Desktop application"を選択
     4. "Authorized redirect URIs"に`http://localhost:4100/code`を追加
     5. Client ID と Client Secret をメモ

## セットアップ手順

1. **リポジトリのクローンとインストール**:

   ```bash
   git clone https://github.com/Yulikepython/gogole-workspace-mcp-server-by-itc.git
   cd google-workspace-mcp-server-by-itc
   npm install
   ```

2. **認証情報の設定**:

   ```bash
   # credentials.json.exampleをコピー
   cp credentials.json.example credentials.json
   ```

   `credentials.json`を以下のように編集:

   ```json
   {
     "web": {
       "client_id": "YOUR_CLIENT_ID",
       "client_secret": "YOUR_CLIENT_SECRET",
       "redirect_uris": ["http://localhost:4100/code"],
       "auth_uri": "https://accounts.google.com/o/oauth2/auth",
       "token_uri": "https://oauth2.googleapis.com/token"
     }
   }
   ```

3. **リフレッシュトークンの取得**:

   ```bash
   node get-refresh-token.js
   ```

   これにより:

   - ブラウザが開き、Google OAuth 認証が実行されます
   - 以下の権限が要求されます:
     - `https://www.googleapis.com/auth/gmail.modify`
     - `https://www.googleapis.com/auth/calendar`
     - `https://www.googleapis.com/auth/gmail.send`
   - 認証情報が`token.json`に保存されます
   - コンソールにリフレッシュトークンが表示されます

4. **MCP 設定の構成**:
   MCP 設定ファイルにサーバー設定を追加:

   - VSCode Claude 拡張機能: `~/Library/Application Support/Code/User/globalStorage/saoudrizwan.claude-dev/settings/cline_mcp_settings.json`
   - Claude デスクトップアプリ: `~/Library/Application Support/Claude/claude_desktop_config.json`

   `mcpServers`オブジェクトに以下を追加:

   ```json
   {
     "mcpServers": {
       "google-workspace": {
         "command": "node",
         "args": ["/path/to/google-workspace-server/build/index.js"],
         "env": {
           "GOOGLE_CLIENT_ID": "your_client_id",
           "GOOGLE_CLIENT_SECRET": "your_client_secret",
           "GOOGLE_REFRESH_TOKEN": "your_refresh_token"
         }
       }
     }
   }
   ```

5. **ビルドと実行**:
   ```bash
   npm run build
   ```

## 使用例

### Gmail 操作

1. **最近のメール一覧**:

   ```json
   {
     "maxResults": 5,
     "query": "is:unread"
   }
   ```

2. **メール検索**:

   ```json
   {
     "query": "from:example@gmail.com has:attachment",
     "maxResults": 10
   }
   ```

3. **メール送信**:

   ```json
   {
     "to": "recipient@example.com",
     "subject": "Hello",
     "body": "Message content",
     "cc": "cc@example.com",
     "bcc": "bcc@example.com"
   }
   ```

4. **メールラベルの変更**:
   ```json
   {
     "id": "message_id",
     "addLabels": ["UNREAD"],
     "removeLabels": ["INBOX"]
   }
   ```

### カレンダー操作

1. **予定一覧**:

   ```json
   {
     "maxResults": 10,
     "timeMin": "2024-01-01T00:00:00Z",
     "timeMax": "2024-12-31T23:59:59Z"
   }
   ```

2. **予定作成**:

   ```json
   {
     "summary": "Team Meeting",
     "location": "Conference Room",
     "description": "Weekly sync-up",
     "start": "2024-01-24T10:00:00Z",
     "end": "2024-01-24T11:00:00Z",
     "attendees": ["colleague@example.com"]
   }
   ```

3. **予定更新**:

   ```json
   {
     "eventId": "event_id",
     "summary": "Updated Meeting Title",
     "location": "Virtual",
     "start": "2024-01-24T11:00:00Z",
     "end": "2024-01-24T12:00:00Z"
   }
   ```

4. **予定削除**:
   ```json
   {
     "eventId": "event_id"
   }
   ```

## トラブルシューティング

1. **認証の問題**:

   - 必要な OAuth スコープが付与されているか確認
   - Client ID と Secret が正しいか確認
   - リフレッシュトークンが有効か確認

2. **API エラー**:
   - Google Cloud Console で API クォータと制限を確認
   - プロジェクトで API が有効化されているか確認
   - リクエストパラメータが正しい形式か確認

## ライセンス

MIT
