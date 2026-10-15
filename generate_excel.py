# -*- coding: utf-8 -*-
"""
Script to generate TIMEOUT_ARCHITECTURE_DESIGN.xlsx
High-precision, corporate-grade architectural workbook
Palette: Monotone / Grayscale
Font: Meiryo UI throughout
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def build_workbook():
    wb = openpyxl.Workbook()
    # Remove default sheet
    default_sheet = wb.active

    # Style constants
    FONT_FAMILY = "Meiryo UI"
    
    # Fonts
    font_title = Font(name=FONT_FAMILY, size=14, bold=True, color="FFFFFF")
    font_section = Font(name=FONT_FAMILY, size=11, bold=True, color="FFFFFF")
    font_tbl_header = Font(name=FONT_FAMILY, size=10, bold=True, color="FFFFFF")
    font_sub_header = Font(name=FONT_FAMILY, size=9, bold=True, color="FFFFFF")
    font_bold = Font(name=FONT_FAMILY, size=9, bold=True, color="1A1A1A")
    font_regular = Font(name=FONT_FAMILY, size=9, bold=False, color="262626")
    font_mono_code = Font(name="Consolas", size=9, bold=False, color="1A1A1A")
    font_alert = Font(name=FONT_FAMILY, size=9, bold=True, color="000000")

    # Fills (Monotone Base)
    fill_title = PatternFill(start_color="1A1A1A", end_color="1A1A1A", fill_type="solid")       # Darkest Charcoal / Black
    fill_section = PatternFill(start_color="333333", end_color="333333", fill_type="solid")     # Dark Charcoal
    fill_tbl_header = PatternFill(start_color="4A4A4A", end_color="4A4A4A", fill_type="solid")  # Medium Dark Gray
    fill_sub_header = PatternFill(start_color="666666", end_color="666666", fill_type="solid")  # Medium Gray
    fill_zebra = PatternFill(start_color="F7F7F7", end_color="F7F7F7", fill_type="solid")       # Very Soft Off-White Gray
    fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")       # Pure White
    fill_accent = PatternFill(start_color="E6E6E6", end_color="E6E6E6", fill_type="solid")      # Light Cool Gray (Highlight)
    fill_alert = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")       # Slate Alert Gray
    
    # Borders
    border_thin_gray = Side(style="thin", color="B3B3B3")
    border_thick_charcoal = Side(style="medium", color="333333")
    border_double_bottom = Side(style="double", color="1A1A1A")
    
    border_all_thin = Border(left=border_thin_gray, right=border_thin_gray, top=border_thin_gray, bottom=border_thin_gray)
    border_header = Border(left=border_thin_gray, right=border_thin_gray, top=border_thick_charcoal, bottom=border_double_bottom)
    border_section = Border(left=border_thick_charcoal, right=border_thick_charcoal, top=border_thick_charcoal, bottom=border_thick_charcoal)

    # Alignments
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_right = Alignment(horizontal="right", vertical="center", wrap_text=True)
    align_title = Alignment(horizontal="left", vertical="center", indent=1)

    # -------------------------------------------------------------
    # SHEET 1: 01_構成概要とタイムアウト基本原則
    # -------------------------------------------------------------
    ws1 = wb.create_sheet(title="01_構成概要と基本原則")
    ws1.views.sheetView[0].showGridLines = True
    
    # Title
    ws1.merge_cells("A1:H1")
    ws1["A1"] = "【システムアーキテクチャ設計標準】マルチ経路ハイブリッド環境におけるタイムアウト設計大全"
    ws1["A1"].font = font_title
    ws1["A1"].fill = fill_title
    ws1["A1"].alignment = align_title
    ws1.row_dimensions[1].height = 36

    ws1.merge_cells("A2:H2")
    ws1["A2"] = "対象環境: AWS API Gateway / ALB / ECS (JBoss EAP 8.1 JTA) / RDS Proxy / Aurora Serverless v2 MySQL (2DB) / バッチ直結"
    ws1["A2"].font = Font(name=FONT_FAMILY, size=9, italic=True, color="595959")
    ws1["A2"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws1.row_dimensions[2].height = 20

    # Section 1
    ws1.merge_cells("A4:H4")
    ws1["A4"] = "1. 対象5経路の通信トポロジーと負荷特性"
    ws1["A4"].font = font_section
    ws1["A4"].fill = fill_section
    ws1["A4"].alignment = align_title
    ws1.row_dimensions[4].height = 26

    headers_s1 = ["経路ID", "通信経路名称", "呼出元 (Client)", "経由コンポーネント (Network / App / DB)", "プロトコル / 連携方式", "負荷特性・トランザクション規模", "タイムアウト上の最大制約・天井", "備考・特性"]
    for col_idx, h in enumerate(headers_s1, start=1):
        cell = ws1.cell(row=5, column=col_idx, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_tbl_header
        cell.alignment = align_center
        cell.border = border_header
    ws1.row_dimensions[5].height = 28

    routes_data = [
        ("経路A", "インターネットAPI", "一般インターネット / モバイルアプリ / Web", "API GW → VPC Link v2 → ALB → ECS (JBoss EAP 8.1 JTA) → RDS Proxy → Aurora v2 (2DB)", "HTTPS (REST/JSON)", "中〜高頻度、低レイテンシ要求、アクセス急変あり", "API Gateway 30秒制限 (HTTP API) / 29秒 (REST)", "不特定多数からのアクセス。Fail-Fastが最重要。"),
        ("経路B", "Salesforce連携API", "Salesforce (Apex Callout / External Services)", "API GW → VPC Link v2 → ALB → ECS (JBoss EAP 8.1 JTA) → RDS Proxy → Aurora v2 (2DB)", "HTTPS (REST/SOAP)", "定常的同期連携、月次・締日のバッチ連携", "Salesforce Apex制約 (最大120秒、標準10秒) + API GW 30秒", "画面からの同期呼出は短時間必須。非同期Apexは長時間可。"),
        ("経路C", "イントラネットAPI", "社内業務システム / 基幹ERP / 拠点端末", "【API GWなし】 → ALB → ECS (JBoss EAP 8.1 JTA) → RDS Proxy → Aurora v2 (2DB)", "HTTPS / HTTP2 (Private)", "【システム最重量・最高負荷】 数万件同期処理、大量データ更新", "API GWなし。ALB idle_timeout (最大4000秒) が上限", "最重要警戒経路。長大JTAがDBロックを保持し他を巻き込むリスク大。"),
        ("経路D", "イントラネット画面 (UI)", "社内ブラウザ (オペレーター端末)", "【API GWなし】 → ALB → ECS (JBoss EAP 8.1 JTA) → RDS Proxy → Aurora v2 (2DB)", "HTTPS (HTML/JSF/Web)", "日中業務時間帯に集中、対話型OLTP", "人間工学的限界 (3〜5秒)。10秒超でリロード・F5連打", "画面操作のため応答性が業務効率に直結。短時間打ち切り必須。"),
        ("経路E", "バッチサーバ (直結)", "バッチサーバ (EC2 / ECS Batch / AWS Batch)", "【API GW / ALB / ECS経由せず】 → RDS Proxy → Aurora v2 (2DB)", "JDBC (MySQL Protocol)", "夜間・日次・月次バルク処理、大量SELECT/UPDATE", "RDS Proxy IdleClientTimeout (1800秒) / Aurora wait_timeout", "RDS Proxyの接続を長時間Pinningし、API/画面の接続を枯渇させる主因。")
    ]

    for r_idx, row_data in enumerate(routes_data, start=6):
        fill = fill_zebra if r_idx % 2 == 0 else fill_white
        for c_idx, val in enumerate(row_data, start=1):
            cell = ws1.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_bold if c_idx in [1, 2] else font_regular
            cell.fill = fill
            cell.border = border_all_thin
            cell.alignment = align_center if c_idx in [1, 5] else align_left
        ws1.row_dimensions[r_idx].height = 42

    # Section 2: Principles
    sec2_row = 12
    ws1.merge_cells(f"A{sec2_row}:H{sec2_row}")
    ws1[f"A{sec2_row}"] = "2. タイムアウト設計の鉄則（ピラミッド階層原則と5大黄金律）"
    ws1[f"A{sec2_row}"].font = font_section
    ws1[f"A{sec2_row}"].fill = fill_section
    ws1[f"A{sec2_row}"].alignment = align_title
    ws1.row_dimensions[sec2_row].height = 26

    principles_headers = ["原則番号", "基本原則名", "基本不等式 / 定義式", "物理的メカニズム・動作原理", "違反時に発生する致命的障害事象"]
    for col_idx, h in enumerate(principles_headers, start=1):
        cell = ws1.cell(row=sec2_row+1, column=col_idx, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_tbl_header
        cell.alignment = align_center
        cell.border = border_header
    ws1.merge_cells(f"E{sec2_row+1}:H{sec2_row+1}")
    ws1.row_dimensions[sec2_row+1].height = 28

    principles_data = [
        ("第1原則", "ピラミッド階層原則 (下流ほど短く)", "Client > API GW > ALB > Undertow > JTA > JDBC > DB Query", "下流（DBクエリ）が最初に諦めてエラーを返せば、JTAがロールバックし、Web層・GW層が順次リソースを解放しながら綺麗なエラーを返却できる。", "上流が先に切断すると、下流で誰も待っていない「ゴースト処理」が動き続け、DBのCPU100%・ロック保持による連鎖崩壊が発生する。"),
        ("第2原則", "Keep-Aliveレースコンディション防止", "JBoss Undertow no-request-timeout > ALB idle_timeout (推奨: +5〜15秒)", "ALBが接続を再利用しようとパケットを送った瞬間に、バックエンドが接続を切断（TCP FIN）すると、ALBは即座にHTTP 502 Bad Gatewayを返す。", "ランダムに「HTTP 502 Bad Gateway」が散発し、再現性が低いため原因究明が極めて困難なインフラ障害となる。"),
        ("第3原則", "2DB JTA分散トランザクション保全", "MySQL innodb_rollback_on_timeout = ON (必須有効化)", "デフォルト(OFF)では、ロック待ちタイムアウト時に「最後に失敗したSQL文のみ」がロールバックされ、トランザクション全体は開いたまま後続が動く。", "2つのAurora DB間で片方だけ更新が進み、重大なデータ不整合・金額ズレが確定発生する。JTAの原子性が完全に破壊される。"),
        ("第4原則", "RDS Proxy Pinning (セッション固定) 隔離", "バッチ接続とオンラインAPI接続のProxy設定・クォータ分離", "XAトランザクション、SET autocommit=0、一時テーブル等を使用すると、Proxyは多重化を停止しDB接続を物理専有（Pinning）する。", "長時間バッチや最重量イントラAPIが全DB接続をPinningし、画面やインターネットAPIが ConnectionBorrowTimeout で全員巻き添え死する。"),
        ("第5原則", "無限ブロックの根絶 (JDBC SocketTimeout)", "MySQL JDBC socketTimeout > 0 (絶対に0/無制限のまま放置しない)", "ネットワーク障害、スイッチ瞬断、Auroraフェイルオーバー時に、OSのTCPソケットがパケットを待ち続けてスレッドが永久ブロックされるのを防ぐ。", "ECSコンテナのワーカースレッドが解放されずに食いつぶされ、タスクがヘルスチェックに失敗して強制再起動・全滅ループに陥る。")
    ]

    for p_idx, p_data in enumerate(principles_data, start=sec2_row+2):
        fill = fill_zebra if p_idx % 2 == 0 else fill_white
        for c_idx, val in enumerate(p_data[:4], start=1):
            cell = ws1.cell(row=p_idx, column=c_idx, value=val)
            cell.font = font_bold if c_idx in [1, 2] else font_regular
            cell.fill = fill
            cell.border = border_all_thin
            cell.alignment = align_center if c_idx in [1] else (align_left if c_idx != 3 else align_center)
        # Merge E to H
        ws1.merge_cells(f"E{p_idx}:H{p_idx}")
        cell_e = ws1.cell(row=p_idx, column=5, value=p_data[4])
        cell_e.font = font_regular
        cell_e.fill = fill
        cell_e.border = border_all_thin
        cell_e.alignment = align_left
        for col_k in range(6, 9):
            ws1.cell(row=p_idx, column=col_k).border = border_all_thin
            ws1.cell(row=p_idx, column=col_k).fill = fill
        ws1.row_dimensions[p_idx].height = 46

    # -------------------------------------------------------------
    # SHEET 2: 02_全レイヤ設定パラメータ一覧
    # -------------------------------------------------------------
    ws2 = wb.create_sheet(title="02_全レイヤ設定パラメータ一覧")
    ws2.views.sheetView[0].showGridLines = True

    # Title
    ws2.merge_cells("A1:I1")
    ws2["A1"] = "【設定カタログ】マルチ経路システムにおける全レイヤ・タイムアウト設定パラメータ完全一覧"
    ws2["A1"].font = font_title
    ws2["A1"].fill = fill_title
    ws2["A1"].alignment = align_title
    ws2.row_dimensions[1].height = 36

    headers_s2 = ["No.", "レイヤ区分", "対象コンポーネント", "設定パラメータ名", "設定ファイル / 設定場所", "デフォルト値", "単位", "役割・動作原理・推奨設定方針", "設定誤り時のリスク・障害影響度"]
    for col_idx, h in enumerate(headers_s2, start=1):
        cell = ws2.cell(row=3, column=col_idx, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_tbl_header
        cell.alignment = align_center
        cell.border = border_header
    ws2.row_dimensions[3].height = 28

    params_catalog = [
        # クライアント
        (1, "クライアント層", "Salesforce Apex", "HttpRequest.setTimeout(timeout_ms)", "Apexコード (Controller / Service)", "10,000", "ms", "外部HTTP呼出のソケット読取待機上限。最大120,000ms。API GW(30s)よりやや長い30,000〜35,000msに設定。", "短すぎると早期切断・不要な再送。長すぎてもAPI GWで30sで切断される。"),
        (2, "クライアント層", "ブラウザ (社内画面UI)", "XMLHttpRequest.timeout / Fetch signal", "画面JavaScript (UIクライアント)", "なし (無限)", "ms", "画面操作のHTTP待機上限。5,000〜10,000msでスピナーを停止しユーザーに案内を出す。", "未設定だと画面が固まり、オペレーターがF5連打で重複リクエスト爆弾を投下する。"),
        # API Gateway
        (3, "関所 (API GW) 層", "API Gateway (HTTP API)", "Integration Timeout", "API Gateway ルート統合設定 (Terraform/Console)", "30,000", "ms", "バックエンド(ALB/ECS)からのレスポンスを待つ絶対上限。30秒固定(緩和不可)。", "バックエンド処理が30秒を超えると即座に504 Gateway Timeoutを返却しゴースト化。"),
        (4, "関所 (API GW) 層", "API Gateway (REST API)", "Integration Timeout", "API Gateway 統合リクエスト設定", "29,000", "ms", "REST APIの場合のデフォルト待機上限。原則29秒。", "同上。バックエンドのJTAトランザクションより必ず大きく保つ必要がある。"),
        # ALB
        (5, "ロードバランサ層", "ALB (属性)", "idle_timeout.timeout_seconds", "ALB 属性設定 (aws_lb.idle_timeout)", "60", "秒", "通信が途絶えたとみなすアイドル時間。最重量イントラAPIに合わせて120〜180秒へ拡張検討。", "短すぎると重いイントラAPIが途中で504切断。長すぎるとゾンビコネクションが残存。"),
        (6, "ロードバランサ層", "ALB ターゲットグループ", "deregistration_delay.timeout_seconds", "ターゲットグループ属性", "300", "秒", "ECSタスクのデプロイ・停止時に既存処理の完了を待つ時間(コネクションドレイン)。", "最重量API処理時間より短いと、デプロイ時に実行中のトランザクションが強制断線。"),
        # JBoss Undertow
        (7, "Webコンテナ層", "JBoss EAP 8.1 Undertow", "listener.no-request-timeout", "standalone.xml / subsystem=undertow", "60,000", "ms", "Keep-Alive接続維持時間。★ALB idle_timeoutより必ず5〜10秒以上長く設定する。", "ALBより短いとKeep-Aliveレースコンディションによる「ランダムHTTP 502」が多発。"),
        (8, "Webコンテナ層", "JBoss EAP 8.1 Undertow", "listener.read-timeout / write-timeout", "standalone.xml / subsystem=undertow", "なし (無制限)", "ms", "リクエスト本体の受信およびレスポンス送信のソケットタイムアウト。", "ネットワーク詰まり時にワーカースレッドが解放されずスレッドプールが枯渇。"),
        # JBoss JTA
        (9, "トランザクション層", "JBoss EAP 8.1 JTA", "coordinator-environment.default-timeout", "standalone.xml / subsystem=transactions", "300", "秒", "JTAトランザクションの全体デフォルト上限。Narayana Reaperが監視。", "300秒のままだと、API GWが切断した後のゾンビ処理が5分間もDBをロックし続ける。"),
        (10, "トランザクション層", "JBoss EAP 8.1 JTA", "@TransactionTimeout (個別指定)", "Javaソースコード (@TransactionTimeout)", "継承", "秒", "メソッド・EJBごとの個別JTAタイムアウト。画面(5s)、外部API(15s)、イントラ(120s)と分離。", "一律設定にすると、画面処理が長引いてフリーズするか、重量APIが即時エラーになる。"),
        # データソース
        (11, "コネクションプール層", "JBoss データソース (IronJacamar)", "pool.blocking-timeout-wait-millis", "standalone.xml / datasources/xa-datasource", "30,000", "ms", "プール内に空きコネクションがない場合にスレッドが待機する時間。3,000〜5,000ms推奨。", "30秒待たせるとDB遅延時にJBossの全スレッドが待機ブロックされ、コンテナが沈没。"),
        (12, "コネクションプール層", "JBoss データソース (IronJacamar)", "pool.idle-timeout-minutes", "standalone.xml / datasources/xa-datasource", "15", "分", "プール内の未使用アイドル接続を破棄する時間。RDS ProxyのIdleClientTimeout未満にする。", "RDS Proxy側が先に切断すると、古い接続を掴んだJBossスレッドで通信エラーが発生。"),
        (13, "コネクションプール層", "JBoss データソース (IronJacamar)", "statement.query-timeout", "standalone.xml / datasources/xa-datasource", "なし (0)", "秒", "発行される全SQL文のデフォルトクエリタイムアウト。JTAタイムアウト未満に設定。", "クエリ単体の暴走を早期に止められない。"),
        # MySQL Connector/J
        (14, "JDBCドライバ層", "MySQL Connector/J", "connectTimeout", "JDBC接続URLパラメータ", "0 (OS依存)", "ms", "RDS ProxyとのTCPハンドシェイク確立待機上限。3,000〜5,000ms推奨。", "Proxy障害時にスレッドがOSのTCPタイムアウト(数分)までブロックされる。"),
        (15, "JDBCドライバ層", "MySQL Connector/J", "socketTimeout", "JDBC接続URLパラメータ", "0 (★無制限)", "ms", "【最重要即死パラメータ】 SQL発行後のソケットRead待機上限。有限値(10,000〜60,000ms)必須。", "「0」のまま放置すると、DBハングや瞬断時にスレッドが永遠に死なず全滅する。"),
        # RDS Proxy
        (16, "DBプロキシ層", "AWS RDS Proxy", "ConnectionBorrowTimeout", "RDS Proxy ターゲットグループ設定", "120", "秒", "Auroraへの接続プールが枯渇した際、クライアントが接続貸出を待つ時間。10〜20秒推奨。", "120秒待たされるとAPI Gateway(30秒)が先に死に、エラーがクライアントに届かない。"),
        (17, "DBプロキシ層", "AWS RDS Proxy", "IdleClientTimeout", "RDS Proxy 設定", "1,800", "秒", "JBossやバッチからの接続がアイドル状態のときProxyが切断する時間。30分。", "JBossのidle-timeout-minutesより短くすると、JBossが古い切断済コネクションを掴む。"),
        (18, "DBプロキシ層", "AWS RDS Proxy", "MaxConnectionsPercent", "RDS Proxy ターゲットグループ設定", "100", "％", "Auroraのmax_connectionsのうちProxyが利用可能な割合。70〜80%推奨。", "100%にするとバッチ直結や管理接続用の枠がなくなり、DB管理操作すら不能になる。"),
        # Aurora MySQL
        (19, "データベース層", "Aurora Serverless v2", "max_execution_time", "DBパラメータグループ", "0 (無制限)", "ms", "SELECT文のオプティマイザ実行時間上限。15,000〜30,000ms推奨。", "事故的な全表スキャンやデカルト積SQLがDBのCPUを100%に張り付かせる。"),
        (20, "データベース層", "Aurora Serverless v2", "innodb_lock_wait_timeout", "DBパラメータグループ", "50", "秒", "行ロック競合時の待機上限。10〜15秒推奨 (重量共存時でも30秒以内)。", "50秒はWeb系には長すぎる。ロック待ち行列が連鎖し、後続リクエストが雪崩式に詰まる。"),
        (21, "データベース層", "Aurora Serverless v2", "innodb_rollback_on_timeout", "DBパラメータグループ", "OFF", "真偽値", "【データ整合性最重要】 ロック待ちタイムアウト時にトランザクション全体を戻すか。", "OFFだとタイムアウトしたSQL文だけが失敗し、後続が動くため2DB間でデータ不整合が確定。必ずON！"),
        (22, "データベース層", "Aurora Serverless v2", "wait_timeout / interactive_timeout", "DBパラメータグループ", "28,800", "秒", "DB側のアイドル接続切断時間(8時間)。RDS Proxy配下ではデフォルトで十分。", "短すぎるとProxyとDB間の接続が不用意に切れ、再接続オーバーヘッドが増加する。")
    ]

    for p_row_idx, p_item in enumerate(params_catalog, start=4):
        fill = fill_zebra if p_row_idx % 2 == 0 else fill_white
        for c_idx, val in enumerate(p_item, start=1):
            cell = ws2.cell(row=p_row_idx, column=c_idx, value=val)
            cell.font = font_bold if c_idx in [1, 2, 4] else font_regular
            cell.fill = fill
            cell.border = border_all_thin
            if c_idx in [1, 2, 6, 7]:
                cell.alignment = align_center
            elif c_idx == 4:
                cell.alignment = align_left
                cell.font = font_mono_code
            else:
                cell.alignment = align_left
        ws2.row_dimensions[p_row_idx].height = 34

    # -------------------------------------------------------------
    # SHEET 3: 03_パターン別タイムアウト推奨値一覧 (Master Matrix)
    # -------------------------------------------------------------
    ws3 = wb.create_sheet(title="03_パターン別タイムアウト推奨値")
    ws3.views.sheetView[0].showGridLines = True

    # Title
    ws3.merge_cells("A1:J1")
    ws3["A1"] = "【推奨値マトリクス】5経路 × 4運用パターン別 タイムアウト設定完全推奨仕様書"
    ws3["A1"].font = font_title
    ws3["A1"].fill = fill_title
    ws3["A1"].alignment = align_title
    ws3.row_dimensions[1].height = 36

    # Pattern description banner
    ws3.merge_cells("A2:J2")
    ws3["A2"] = "P1: 標準OLTPバランス型（王道設計） | P2: イントラ超重量・バッチ共存型（現行要件対応） | P3: 超耐障害・フェイルファスト型（金融・EC型） | P4: アーキテクチャ非同期分離型（将来推奨）"
    ws3["A2"].font = Font(name=FONT_FAMILY, size=9, bold=True, color="333333")
    ws3["A2"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws3.row_dimensions[2].height = 22

    headers_s3 = [
        "No.", "レイヤ区分", "設定パラメータ名", "単位", "対象経路",
        "パターン1: 標準バランス型\n[推奨値 / 動作イメージ]",
        "パターン2: イントラ重量・共存型\n[推奨値 / 動作イメージ]",
        "パターン3: フェイルファスト型\n[推奨値 / 動作イメージ]",
        "パターン4: 非同期近代化推奨型\n[推奨値 / 動作イメージ]",
        "設計根拠と留意事項"
    ]
    for col_idx, h in enumerate(headers_s3, start=1):
        cell = ws3.cell(row=4, column=col_idx, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_tbl_header
        cell.alignment = align_center
        cell.border = border_header
    ws3.row_dimensions[4].height = 38

    matrix_data = [
        # クライアント
        (1, "クライアント層", "Salesforce Apex Callout", "ms", "経路B",
         "30,000 ms\n(API GW 30sに同期)", "35,000 ms\n(API GW切断を確実に検知)", "15,000 ms\n(迅速に失敗検知)", "10,000 ms\n(非同期受付即答連携)",
         "SalesforceのHttpRequest.setTimeout上限は120sだが、API GWの30s制約に従う。"),
        (2, "クライアント層", "社内画面UI Timeout", "ms", "経路D",
         "10,000 ms\n(操作待ち限界)", "15,000 ms\n(やや長めの検索を許容)", "5,000 ms\n(俊敏な操作感)", "5,000 ms\n(重い処理は非同期化)",
         "オペレーターが我慢できる限界。これを超えると再読み込みやF5連打が起きる。"),
        # API Gateway
        (3, "API Gateway層", "Integration Timeout", "ms", "経路A, B",
         "30,000 ms (HTTP) / 29,000 ms (REST)", "30,000 ms (HTTP) / 29,000 ms (REST)", "15,000 ms\n(手前で早期足切り)", "15,000 ms\n(全同期処理を短縮)",
         "HTTP APIは30秒固定。REST APIはデフォルト29秒。JBossのJTAより必ず長くする。"),
        # ALB
        (4, "ALB層", "idle_timeout.timeout_seconds", "秒", "全経路",
         "60 秒\n(標準デフォルト)", "180 秒\n(★重量イントラAPIを通過させる)", "30 秒\n(ゾンビコネクション早期刈取)", "60 秒\n(標準に戻す)",
         "P2では最重量API(120s)を落とさないため180sに拡大。Undertow keep-aliveとの関係に厳重注意。"),
        (5, "ALB層", "deregistration_delay.timeout_seconds", "秒", "全経路",
         "60 秒\n(標準)", "200 秒\n(デプロイ時処理保護)", "30 秒\n(高速タスク入替)", "60 秒\n(標準)",
         "ECSタスク更新時に実行中のトランザクションが切断されないよう、最大処理時間以上に設定。"),
        # Undertow
        (6, "JBoss Undertow層", "listener.no-request-timeout", "ms", "全経路",
         "65,000 ms\n(ALB 60s + 5s)", "190,000 ms\n(★ALB 180s + 10s)", "35,000 ms\n(ALB 30s + 5s)", "65,000 ms\n(ALB 60s + 5s)",
         "【第2原則】ALB idle_timeoutより必ず長くする。短いとランダムHTTP 502が発生する。"),
        # JBoss JTA
        (7, "JTA Narayana層", "coordinator default-timeout", "秒", "共通",
         "30 秒\n(安全な全体デフォルト)", "150 秒\n(重量処理用デフォルト)", "15 秒\n(厳格なデフォルト)", "30 秒\n(標準デフォルト)",
         "JBoss全体のフォールバック値。デフォルト300秒の放置は絶対に禁止。"),
        (8, "JTA Narayana層", "@TransactionTimeout (経路A: ネットAPI)", "秒", "経路A",
         "15 秒\n(API GW手前で撤退)", "20 秒\n(API GW手前で撤退)", "8 秒\n(超高速フェイルファスト)", "10 秒\n(短時間OLTP)",
         "API Gateway(30s)より手前でJTAがロールバックし、リソースを速やかに解放する。"),
        (9, "JTA Narayana層", "@TransactionTimeout (経路B: Salesforce)", "秒", "経路B",
         "15 秒\n(API GW手前で撤退)", "20 秒\n(API GW手前で撤退)", "8 秒\n(超高速フェイルファスト)", "10 秒\n(短時間OLTP)",
         "Salesforce側のタイムアウトより前にJBossから503/400系で整然とエラーを戻す。"),
        (10, "JTA Narayana層", "@TransactionTimeout (経路C: イントラ重量API)", "秒", "経路C",
         "20 秒\n(重い処理も20秒制限)", "120 秒\n(★最重量処理を完走させる)", "12 秒\n(重量処理も早期制限)", "5 秒 (受付即答) / Worker 300秒",
         "P2の核心設定。API GWを通らないため120sを許容。P4では非同期Workerに逃がす。"),
        (11, "JTA Narayana層", "@TransactionTimeout (経路D: イントラ画面)", "秒", "経路D",
         "5 秒\n(快適な画面レスポンス)", "10 秒\n(最大我慢時間)", "3 秒\n(俊敏性重視)", "3 秒\n(俊敏性重視)",
         "画面処理は5秒以内。遅い場合は集計クエリのインデックス不備かロック競合。"),
        # データソース
        (12, "コネクションプール層", "pool.blocking-timeout-wait-millis", "ms", "ECS",
         "5,000 ms\n(5秒待ってダメなら諦める)", "10,000 ms\n(重量APIのコネクション待ち許容)", "3,000 ms\n(待たずに即座にエラー)", "3,000 ms\n(プール健全性維持)",
         "DB混雑時にJBossのスレッドが芋づる式にブロックされるのを防ぐ防壁。"),
        # JDBCドライバ
        (13, "MySQL Connector/J", "connectTimeout", "ms", "全経路",
         "5,000 ms\n(迅速なハンドシェイク)", "5,000 ms\n(迅速なハンドシェイク)", "3,000 ms\n(高速エラー検知)", "5,000 ms\n(迅速なハンドシェイク)",
         "RDS ProxyへのTCP接続確立待機。正常時は数msで繋がるため5秒で十分。"),
        (14, "MySQL Connector/J", "socketTimeout (経路A, B, D: オンライン)", "ms", "経路A,B,D",
         "20,000 ms\n(JTAタイムアウト直後)", "25,000 ms\n(JTAタイムアウト直後)", "10,000 ms\n(JTAタイムアウト直後)", "15,000 ms\n(JTAタイムアウト直後)",
         "【第5原則】絶対に0にしない。JTAトランザクション時間より少し長めに設定。"),
        (15, "MySQL Connector/J", "socketTimeout (経路C: イントラ重量API)", "ms", "経路C",
         "25,000 ms\n(オンライン同期)", "130,000 ms\n(★120sのJTAを保護)", "15,000 ms\n(短時間打ち切り)", "10,000 ms (受付) / 320,000 ms (Worker)",
         "P2では120秒のJTAトランザクションを途中で切断しないよう130秒に設定。"),
        (16, "MySQL Connector/J", "socketTimeout (経路E: バッチ直結)", "ms", "経路E",
         "60,000 ms\n(1クエリ1分以内)", "300,000 ms\n(長大クエリ5分許容)", "60,000 ms\n(短時間チャンク推奨)", "600,000 ms (専用Proxy経由)",
         "バッチが直接RDS Proxyに繋ぐため、バッチ専用の長めのSocketTimeoutを設定。"),
        # RDS Proxy
        (17, "RDS Proxy層", "ConnectionBorrowTimeout", "秒", "共通",
         "15 秒\n(API GWの半値)", "30 秒\n(重量処理待ちバッファ)", "5 秒\n(接続待ちせずフェイルファスト)", "10 秒\n(待機短縮)",
         "デフォルト120秒の短縮必須。これを超えるとAPI GWが先に死んでカスケード障害になる。"),
        (18, "RDS Proxy層", "MaxConnectionsPercent", "％", "共通",
         "80 %\n(バッチ・管理用20%予約)", "70 %\n(バッチ用30%予約)", "80 %\n(バッチ・管理用20%予約)", "ECS用Proxy: 70% / バッチ用Proxy: 20%",
         "ECS経由とバッチ直結が同居するため、ECS側Proxyの接続枠を70〜80%に絞る。"),
        # Aurora MySQL
        (19, "Aurora MySQL層", "max_execution_time", "ms", "オンライン",
         "15,000 ms\n(暴走SELECT防止)", "25,000 ms\n(重量SELECT許容)", "8,000 ms\n(即時打ち切り)", "10,000 ms\n(短時間保証)",
         "SELECT文の強制打切り時間。バッチ処理のセッションでは無効(0)化する。"),
        (20, "Aurora MySQL層", "innodb_lock_wait_timeout", "秒", "共通",
         "10 秒\n(ロック待ち行列防止)", "30 秒\n(重量更新のロック待ち許容)", "5 秒\n(ミリ秒単位でロック解放)", "10 秒\n(健全な短時間ロック)",
         "行ロック待ちの上限。デフォルト50秒から大幅短縮し、他経路の巻き添えを防ぐ。"),
        (21, "Aurora MySQL層", "innodb_rollback_on_timeout", "真偽値", "共通",
         "ON (必須)\n[全体ロールバック]", "ON (必須)\n[全体ロールバック]", "ON (必須)\n[全体ロールバック]", "ON (必須)\n[全体ロールバック]",
         "【第3原則・最重要】ロック待ちタイムアウト時に2DBのトランザクションを完全破棄。OFFは絶対禁止！")
    ]

    for m_row_idx, m_item in enumerate(matrix_data, start=5):
        fill = fill_zebra if m_row_idx % 2 == 0 else fill_white
        for c_idx, val in enumerate(m_item, start=1):
            cell = ws3.cell(row=m_row_idx, column=c_idx, value=val)
            cell.font = font_bold if c_idx in [1, 2, 3] else font_regular
            cell.fill = fill
            cell.border = border_all_thin
            if c_idx in [1, 4, 5]:
                cell.alignment = align_center
            elif c_idx in [6, 7, 8, 9]:
                cell.alignment = align_center
                if "★" in str(val):
                    cell.fill = fill_accent
                    cell.font = font_alert
            else:
                cell.alignment = align_left
        ws3.row_dimensions[m_row_idx].height = 42

    # -------------------------------------------------------------
    # SHEET 4: 04_経路別・障害シーケンス詳細
    # -------------------------------------------------------------
    ws4 = wb.create_sheet(title="04_経路別・障害シーケンス詳細")
    ws4.views.sheetView[0].showGridLines = True

    # Title
    ws4.merge_cells("A1:I1")
    ws4["A1"] = "【障害分析・根本原因診断】4大障害シナリオの動作メカニズムとタイムアウト連鎖抑止設計"
    ws4["A1"].font = font_title
    ws4["A1"].fill = fill_title
    ws4["A1"].alignment = align_title
    ws4.row_dimensions[1].height = 36

    headers_s4 = ["シナリオNo.", "障害事象名称", "主な発生箇所", "発生メカニズム・時系列シーケンス (Timeline)", "根本原因 (Root Cause)", "タイムアウトによる防止策 (設定効果)", "アプリ・インフラ補完対策", "復旧動作イメージ", "影響を受ける経路"]
    for col_idx, h in enumerate(headers_s4, start=1):
        cell = ws4.cell(row=3, column=col_idx, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_tbl_header
        cell.alignment = align_center
        cell.border = border_header
    ws4.row_dimensions[3].height = 28

    scenarios_data = [
        ("シナリオ1", "最重量イントラAPIによる長時間ロック保持とカスケード全滅", "経路C (イントラAPI) ──► Aurora DB1/DB2",
         "[T+0s] イントラAPIが数万件の洗替えSQL(UPDATE)を開始し行ロックを大量保持。\n"
         "[T+10s] 社内画面(経路D)とSalesforce(経路B)が同テーブルの更新を要求しLock Waitに並ぶ。\n"
         "[T+30s] API GWが30秒を迎えて504タイムアウト返却。Salesforceはリトライ開始。\n"
         "[T+50s] 未チューニングのinnodb_lock_wait_timeout(50s)でエラー発生するがOFFのため不整合。\n"
         "[T+60s] JBossの全コネクションがロック待ちで埋まり、ネットAPI(経路A)も巻き添え停止。",
         "最重量APIのJTAトランザクション時間とAuroraのinnodb_lock_wait_timeoutが長すぎ、ロック解放前に上流が切断・リトライして雪崩を起こした。",
         "① innodb_lock_wait_timeout を10秒に短縮。\n"
         "② innodb_rollback_on_timeout = ON で中途半端なロックを即座に消去。\n"
         "③ 外部APIのJTAタイムアウトを15秒にして早期離脱。",
         "最重量APIのバッチ処理化（Chunk分割コミット）、または更新対象テーブルのパーティショニングと更新キーの分散。",
         "ロック競合が10秒で打ち切られ、JBossが速やかにロールバック。JBossスレッドが即座に解放され、他経路への汚染が食い止められる。",
         "全経路 (特にB, D, A)"),

        ("シナリオ2", "バッチ処理によるRDS Proxy Pinningと接続枯渇", "経路E (バッチ直結) ──► RDS Proxy",
         "[T+0s] バッチがautocommit=falseで大量バルクINSERTを開始。\n"
         "[T+5s] RDS Proxyがトランザクションを検知し、DB接続を専用固定（Pinning）。\n"
         "[T+10s] バッチが並行で複数スレッド走り、Proxyの全Target接続をPinningで使い切る。\n"
         "[T+15s] ECSからAPIや画面のリクエストが来るが、Proxyに空き接続がない。\n"
         "[T+45s] ConnectionBorrowTimeout(デフォルト120s)の間、全リクエストが待たされる。\n"
         "[T+30s] API GatewayとALBが先にタイムアウトし、画面と外部APIが全滅。",
         "バッチ直結処理がRDS Proxyの多重化を無効化（Pinning）し、オンライン用の接続プールを食い尽くした。",
         "① RDS Proxyの ConnectionBorrowTimeout を15秒に短縮。\n"
         "② JBossの blocking-timeout-wait-millis を5000msに設定。\n"
         "③ バッチ側のクエリ・SocketTimeoutを適切に絞る。",
         "バッチ専用のRDS Proxyエンドポイント（またはAurora Reader直結）を用意し、オンライン用Proxyプールと物理的に分離する。",
         "Proxy待ちが15秒で安全にタイムアウトし、JBossスレッドの無限蓄積を防止。バッチのコミット間隔を1,000件単位に細分化。",
         "経路A, 経路B, 経路C, 経路D"),

        ("シナリオ3", "Undertow Keep-Alive と ALB Idle Timeout の競合 (ランダム502)", "ALB ◄──► ECS JBoss Undertow",
         "[T+0s] クライアントからのリクエストが完了し、ALBとJBoss間のTCP接続がアイドル状態になる。\n"
         "[T+60.000s] JBoss UndertowがKeep-Alive時間(60秒)に達したためTCP FINパケットを送信。\n"
         "[T+60.001s] ちょうど同じ瞬間、ALBが新しいリクエストをその接続へ転送。\n"
         "[T+60.002s] 送信パケットが交差（クロス）し、ALBにTCP RSTが返る。\n"
         "[T+60.003s] ALBはバックエンド異常と判定し、クライアントへ「HTTP 502 Bad Gateway」を返す。",
         "JBoss UndertowのKeep-Alive時間(no-request-timeout)が、ALBの idle_timeout 以下になっていたことによるレースコンディション。",
         "Undertowの no-request-timeout を ALBの idle_timeout よりも必ず5〜10秒以上長く設定する。\n"
         "例: ALB 60秒 ──► Undertow 65〜70秒\n"
         "例: ALB 180秒 ──► Undertow 190秒",
         "インフラIaC（Terraform）とJBoss standalone.xmlの設定値をCI/CDパイプラインで自動整合性チェックする。",
         "JBoss側から勝手に接続を切断しなくなり、ALBが主導権を持って安全にコネクションをクローズするため、502エラーが完全にゼロになる。",
         "経路A, 経路B, 経路C, 経路D"),

        ("シナリオ4", "Aurora Serverless v2 のACUスケール追従遅延と接続過負荷", "ECS / バッチ ──► Aurora Serverless v2",
         "[T+0s] 突発的なアクセス増または最重量イントラAPIの開始により、DB CPU使用率が95%に急騰。\n"
         "[T+3s] Aurora Serverless v2がスケールアップ（例: 2 ACU ──► 16 ACU）を開始。\n"
         "[T+8s] スケーリング完了までの数秒間、クエリ処理速度が一時的に低下。\n"
         "[T+12s] クエリ遅延に伴い、JBossとバッチからの接続要求が急増しキューイング。\n"
         "[T+25s] 適切な max_execution_time がないため、遅延クエリが雪だるま式に蓄積。",
         "Serverless v2は極めて高速にスケールするが、ゼロ秒でスケールするわけではない。スケール途中の過負荷時にクエリを野放しにしたこと。",
         "① Aurora max_execution_time を15,000msに設定し、長引くSELECTを自動強制終了。\n"
         "② JTAトランザクションタイムアウトにより、過負荷時の処理を迅速に間引く。",
         "Aurora Serverless v2の最小ACU（Min ACU）をピーク予測に合わせてあらかじめ底上げ（プロビジョンド的運用）しておく。",
         "スケーリング中の数秒間に重いクエリが安全に遮断され、DBの完全窒息を回避。ACUがスケールアップした後に安定稼働へ復帰。",
         "全経路")
    ]

    for s_row_idx, s_item in enumerate(scenarios_data, start=4):
        fill = fill_zebra if s_row_idx % 2 == 0 else fill_white
        for c_idx, val in enumerate(s_item, start=1):
            cell = ws4.cell(row=s_row_idx, column=c_idx, value=val)
            cell.font = font_bold if c_idx in [1, 2] else font_regular
            cell.fill = fill
            cell.border = border_all_thin
            if c_idx in [1, 3, 9]:
                cell.alignment = align_center
            else:
                cell.alignment = align_left
        ws4.row_dimensions[s_row_idx].height = 85

    # -------------------------------------------------------------
    # SHEET 5: 05_設計・運用チェックリスト
    # -------------------------------------------------------------
    ws5 = wb.create_sheet(title="05_設計・運用チェックリスト")
    ws5.views.sheetView[0].showGridLines = True

    # Title
    ws5.merge_cells("A1:H1")
    ws5["A1"] = "【本番導入・運用監視チェックリスト】タイムアウト設計検証シート & CloudWatch監視基準"
    ws5["A1"].font = font_title
    ws5["A1"].fill = fill_title
    ws5["A1"].alignment = align_title
    ws5.row_dimensions[1].height = 36

    # Section 1: Pre-launch Verification Checklist
    ws5.merge_cells("A3:H3")
    ws5["A3"] = "1. システム本番稼働前・アーキテクチャ整合性チェックリスト"
    ws5["A3"].font = font_section
    ws5["A3"].fill = fill_section
    ws5["A3"].alignment = align_title
    ws5.row_dimensions[3].height = 26

    headers_chk = ["No.", "検証カテゴリ", "検査項目 / 設定パラメータ", "必須判定基準 (Pass Criteria)", "重要度 / 影響度", "確認対象設定ファイル / 画面", "検証方法・コマンド", "確認結果"]
    for col_idx, h in enumerate(headers_chk, start=1):
        cell = ws5.cell(row=4, column=col_idx, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_tbl_header
        cell.alignment = align_center
        cell.border = border_header
    ws5.row_dimensions[4].height = 28

    checklist_items = [
        (1, "Web/ALB整合性", "Undertow vs ALB Keep-Alive", "Undertow no-request-timeout > ALB idle_timeout (+5s以上)", "CRITICAL (最重要)", "standalone.xml & Terraform aws_lb", "設定値の直接比較検証", "[  ] PASS"),
        (2, "デプロイ保全", "ALB Deregistration Delay", "deregistration_delay >= 最重量イントラAPI最大時間 (P2なら200s以上)", "HIGH (高)", "Terraform aws_lb_target_group", "ターゲットグループ属性の確認", "[  ] PASS"),
        (3, "API GW防壁", "API Gateway Integration Timeout", "HTTP API: 30s固定 / REST: 29s。JBoss JTA(15〜20s)より大", "CRITICAL (最重要)", "API Gateway コンソール / IaC", "インテグレーション設定値の確認", "[  ] PASS"),
        (4, "JTAデフォルト", "JBoss Narayana デフォルト時間", "coordinator default-timeout が300秒から短縮されていること", "HIGH (高)", "standalone.xml (transactions subsystem)", "CLI: /subsystem=transactions:read-resource", "[  ] PASS"),
        (5, "メソッド別JTA", "@TransactionTimeout 設定", "画面(5s)、外部API(15s)、最重量API(120s)が個別指定されていること", "HIGH (高)", "Javaソースコード / アノテーション", "ソースコード静的解析 / Grep", "[  ] PASS"),
        (6, "プール枯渇防止", "データソース blocking-timeout", "blocking-timeout-wait-millis が 3000〜5000ms に設定されていること", "HIGH (高)", "standalone.xml (datasources subsystem)", "CLI: /subsystem=datasources:read-resource", "[  ] PASS"),
        (7, "無限待ち撲滅", "JDBC socketTimeout 設定", "socketTimeout が「0(無制限)」ではなく有限値(20000ms等)であること", "CRITICAL (最重要)", "standalone.xml (接続URL)", "JDBC接続URL文字列の目視確認", "[  ] PASS"),
        (8, "Proxy接続待機", "RDS Proxy BorrowTimeout", "ConnectionBorrowTimeout が 120s から 15〜30s に短縮されていること", "HIGH (高)", "RDS Proxy ターゲットグループ設定", "AWS CLI: describe-db-proxy-target-groups", "[  ] PASS"),
        (9, "データ整合性", "MySQL innodb_rollback_on_timeout", "DBパラメータグループで確実に「ON」に設定されていること", "CRITICAL (最重要)", "Aurora DBパラメータグループ", "SQL: SHOW VARIABLES LIKE 'innodb_rollback_on_timeout';", "[  ] PASS"),
        (10, "ロック待機適正", "MySQL innodb_lock_wait_timeout", "デフォルト50秒から 10〜30秒 に短縮されていること", "HIGH (高)", "Aurora DBパラメータグループ", "SQL: SHOW VARIABLES LIKE 'innodb_lock_wait_timeout';", "[  ] PASS"),
        (11, "暴走SELECT防止", "MySQL max_execution_time", "オンライン処理用セッションで 15000〜25000ms に設定されていること", "MEDIUM (中)", "DBパラメータグループ / init_connect", "SQL: SHOW VARIABLES LIKE 'max_execution_time';", "[  ] PASS"),
        (12, "バッチ枠予約", "RDS Proxy MaxConnectionsPercent", "オンライン用Proxyの最大接続割合が 70〜80% に制限されていること", "HIGH (高)", "RDS Proxy ターゲットグループ設定", "AWS CLI: describe-db-proxy-target-groups", "[  ] PASS")
    ]

    for c_row_idx, c_item in enumerate(checklist_items, start=5):
        fill = fill_zebra if c_row_idx % 2 == 0 else fill_white
        for c_idx, val in enumerate(c_item, start=1):
            cell = ws5.cell(row=c_row_idx, column=c_idx, value=val)
            cell.font = font_bold if c_idx in [1, 3, 5] else font_regular
            cell.fill = fill
            cell.border = border_all_thin
            if c_idx in [1, 2, 5, 8]:
                cell.alignment = align_center
                if "CRITICAL" in str(val):
                    cell.fill = fill_alert
                    cell.font = font_alert
            else:
                cell.alignment = align_left
        ws5.row_dimensions[c_row_idx].height = 28

    # Section 2: CloudWatch Monitoring Guide
    sec2_cw_row = 18
    ws5.merge_cells(f"A{sec2_cw_row}:H{sec2_cw_row}")
    ws5[f"A{sec2_cw_row}"] = "2. 本番運用時 CloudWatch 監視メトリクスと推奨アラーム閾値一覧"
    ws5[f"A{sec2_cw_row}"].font = font_section
    ws5[f"A{sec2_cw_row}"].fill = fill_section
    ws5[f"A{sec2_cw_row}"].alignment = align_title
    ws5.row_dimensions[sec2_cw_row].height = 26

    headers_cw = ["No.", "監視対象リソース", "CloudWatch Namespace", "メトリクス名", "警戒アラーム閾値 (Warning)", "危険アラーム閾値 (Critical)", "アラーム検知時の意味・障害予兆", "初動アクション・推奨対応"]
    for col_idx, h in enumerate(headers_cw, start=1):
        cell = ws5.cell(row=sec2_cw_row+1, column=col_idx, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_tbl_header
        cell.alignment = align_center
        cell.border = border_header
    ws5.row_dimensions[sec2_cw_row+1].height = 28

    cw_metrics = [
        (1, "RDS Proxy", "AWS/RDSProxy", "DatabaseConnectionsCurrentlyBorrowed", "全DB接続の 70% 超 (3分継続)", "全DB接続の 85% 超 (1分継続)", "接続プール枯渇予兆。バッチや重量APIのPinning急増。", "実行中トランザクションの特定、バッチ一時停止。"),
        (2, "RDS Proxy", "AWS/RDSProxy", "QueryRequestLatency", "平均 > 500 ms", "平均 > 2,000 ms", "Proxy内部での借用待ち遅延、Auroraスケーリング遅延。", "Aurora ACU使用率確認、スロークエリ特定。"),
        (3, "ALB", "AWS/ApplicationELB", "HTTPCode_Target_5XX_Count", "合計 > 10 回 / 5分", "合計 > 50 回 / 1分", "JBossコンテナ内でのタイムアウト、エラー噴出。", "JBoss server.logの調査、JTAロールバック確認。"),
        (4, "ALB", "AWS/ApplicationELB", "HTTPCode_ELB_504_Count", "合計 > 5 回 / 5分", "合計 > 20 回 / 1分", "ALB idle_timeout超過。最重量APIの遅延暴走。", "イントラAPIの実行状況確認、DBロック状況調査。"),
        (5, "ALB", "AWS/ApplicationELB", "HTTPCode_ELB_502_Count", "合計 > 1 回 / 5分", "合計 > 10 回 / 1分", "Keep-Aliveレースコンディション、またはコンテナ異常終了。", "Undertow no-request-timeout設定値の再検証。"),
        (6, "Aurora MySQL", "AWS/RDS", "ServerlessDatabaseCapacity", "最大ACUの 80% 超", "最大ACUに到達 (100%)", "Auroraの処理能力が上限到達。これ以上のスケール不可。", "不要な重クエリのキル、最大ACU設定の引き上げ検討。"),
        (7, "Aurora MySQL", "AWS/RDS", "BlockedTransactions", "現在値 > 5 件", "現在値 > 20 件", "深刻な行ロック競合。カスケード全滅の初期兆候。", "SHOW ENGINE INNODB STATUS でロック元SQLを即座に特定。"),
        (8, "API Gateway", "AWS/ApiGateway", "5XXError", "エラー率 > 1 %", "エラー率 > 5 %", "API GWでの30秒タイムアウト(504)多発。", "経路A・Bのバックエンド処理時間急増を調査。")
    ]

    for cw_row_idx, cw_item in enumerate(cw_metrics, start=sec2_cw_row+2):
        fill = fill_zebra if cw_row_idx % 2 == 0 else fill_white
        for c_idx, val in enumerate(cw_item, start=1):
            cell = ws5.cell(row=cw_row_idx, column=c_idx, value=val)
            cell.font = font_bold if c_idx in [1, 4] else font_regular
            cell.fill = fill
            cell.border = border_all_thin
            if c_idx in [1, 2, 3]:
                cell.alignment = align_center
            elif c_idx in [5, 6]:
                cell.alignment = align_center
                cell.font = font_bold
            else:
                cell.alignment = align_left
        ws5.row_dimensions[cw_row_idx].height = 32

    # -------------------------------------------------------------
    # AUTO-FIT COLUMN WIDTHS ACROSS ALL SHEETS (with safety buffers)
    # -------------------------------------------------------------
    sheet_col_widths = {
        "01_構成概要と基本原則": {1: 10, 2: 24, 3: 26, 4: 45, 5: 22, 6: 35, 7: 35, 8: 35},
        "02_全レイヤ設定パラメータ一覧": {1: 8, 2: 18, 3: 26, 4: 36, 5: 35, 6: 16, 7: 10, 8: 52, 9: 50},
        "03_パターン別タイムアウト推奨値": {1: 6, 2: 18, 3: 32, 4: 8, 5: 14, 6: 28, 7: 30, 8: 28, 9: 30, 10: 42},
        "04_経路別・障害シーケンス詳細": {1: 12, 2: 30, 3: 28, 4: 55, 5: 38, 6: 42, 7: 42, 8: 45, 9: 20},
        "05_設計・運用チェックリスト": {1: 6, 2: 18, 3: 32, 4: 45, 5: 18, 6: 35, 7: 35, 8: 14}
    }

    for s_name, widths in sheet_col_widths.items():
        if s_name in wb.sheetnames:
            target_ws = wb[s_name]
            for col_i, w in widths.items():
                col_letter = get_column_letter(col_i)
                target_ws.column_dimensions[col_letter].width = w

    # Remove default sheet if exists
    if "Sheet" in wb.sheetnames and len(wb.sheetnames) > 1:
        wb.remove(wb["Sheet"])

    # Save
    output_filename = "TIMEOUT_ARCHITECTURE_DESIGN.xlsx"
    wb.save(output_filename)
    print(f"Successfully generated {output_filename}")

if __name__ == "__main__":
    build_workbook()
