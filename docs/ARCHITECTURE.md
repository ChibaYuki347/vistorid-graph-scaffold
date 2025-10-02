
# VisitorID — Add-in × Graph Integration (Delta-only 基本構成)

> 目的：本文依存をやめ、**Compose中にID確定 → 非表示のカスタムデータ保存 → バックエンドでDB正規化**。同期は **5分 Delta**（必要に応じて通知＋Delta）。

## 責務分担（誰が何をするか）

| 層 | 責務 |
|---|---|
| **Outlook アドイン（クライアント）** | 1) `saveAsync` で itemId を確定<br>2) `CustomProperties` に `visitorId` を保存<br>3) (任意) 本文に案内行を追記（人向け）<br>4) **バックエンドへ POST**（`tenant, room, itemId(ews), itemId(rest?)` など） |
| **バックエンド API** | 1) `translateExchangeIds` で Graph用の **restId/immutableId** へ変換<br>2) (任意) `iCalUId` を取得して正規化キーを確定<br>3) **DB に `{tenant, room, graphEventId, iCalUId, visitorId}` をUPSERT**<br>4) (任意) イベントへ **Open extension** を書き込み |
| **同期ワーカー（5分 Delta）** | 1) 会議室の `calendarView/delta`（昨日〜+7日）<br>2) 受け取ったイベントの `id/iCalUId` で **DB JOIN** → Visitor情報を付与してキャッシュ更新 |

> 時刻整合：`Prefer: outlook.timezone="Tokyo Standard Time"` を使用。本文不要のため `Prefer: outlook.body-content-type="text"` は省略可。

## シーケンス（Compose → 保存 → 正規化）

```mermaid
sequenceDiagram
  autonumber
  participant U as ユーザー
  participant A as Outlook アドイン
  participant O as Exchange/Outlook
  participant B as 自社バックエンド(API)
  participant G as Microsoft Graph
  participant D as DB

  U->>A: 来訪者情報を入力（VisitorID発行）
  A->>O: saveAsync() で予定保存（招待は未送信）
  O-->>A: itemId（EWS形式）を返す
  A->>A: CustomProperties.set("visitorId", ...).saveAsync()
  A->>B: POST /link {tenant, room, ewsItemId, restItemId? , visitorId}
  B->>G: translateExchangeIds(ews/rest -> restImmutableEntryId)
  G-->>B: graphEventId(immutable)
  B->>G: GET /events/{id}?$select=iCalUId,subject,start,end
  G-->>B: iCalUId 他
  B->>D: UPSERT {tenant, room, graphEventId, iCalUId, visitorId}
```

## シーケンス（5分 Delta → キャッシュ更新）

```mermaid
sequenceDiagram
  autonumber
  participant W as Delta Worker (5分)
  participant G as Microsoft Graph
  participant D as DB/Cache
  participant UI as 会議室一覧UI

  W->>G: GET /users/{room}/calendarView/delta?start=Y-1&end=+7
  G-->>W: 変更済みイベント群 + next/deltaLink
  W->>D: JOIN(iCalUId/graphId) で Visitor を付与→キャッシュ更新
  UI->>D: 直近の予定を読み込み
  D-->>UI: 表示用データ（P95≦10秒を狙う）
```

## スロットリング/並列度
- ルーム毎に直列処理、テナント別にキュー/ワーカーを分離（相互影響を遮断）。
- `$batch` は **20/回 上限**を厳守。429 は `Retry-After` 準拠。

## 注意
- `calendarView/delta` は `?$select/$expand/$filter/$orderby/$search` **非対応**。拡張値は追加 GET または DB 参照で解決。
