# AlarmMasterAutomation 運用指示書（Claude Code専用・統合版）

このファイルは `C:\work\AlarmMasterAutomation\` 直下に置く。
Claude Codeをこのフォルダで起動すると自動的に読み込まれる。
claude.ai（チャット版）を介さず、このフォルダ内で開発〜反映〜検証を完結させることを目的とする。

## プロジェクト概要

au（KDDI）5GC「10.8M Expansion」のアラーム監視業務を効率化するExcel VBAツール群。
CSV集約 → 事前精査テキスト生成 → 台帳登録 → 差分比較、が基本パイプライン。
既存動作を壊さないことを新機能追加より最優先する。

## 参照必須ドキュメント（作業前に必ず読むこと）

- `docs\vba-alarm-tooling-SKILL.md` … コーディング規約・命名規則・設計方針の正本。
  「VBAで～」「.basを追加／修正」等の依頼を受けたら、必ず先にこれを読んでから着手する。
- `src\` 配下 … 現行の各モジュールのマスターコピー。修正前に必ず対象モジュールと、
  影響を受けそうな他モジュールを読んで影響範囲を確認する。
  **2026-07-31改訂**: `src\`直下は`Alarm_Master.xlsm`向け、`src\Alarm_Master\`と
  `src\PERSONAL\`はそれぞれの対象ブック専用のサブフォルダ（詳細は次項）。

## このディレクトリの役割

**対象ブックは2つある**（2026-07-31実機確認・PERSONAL.XLSBを新たに管理対象へ追加）:

1. `Alarm_Master.xlsm`（一元管理台帳） — `src\`直下および`src\Alarm_Master\`のモジュール
2. `PERSONAL.XLSB`（個人用マクロブック。`C:\Users\kinoue\AppData\Roaming\Microsoft\Excel\XLSTART\PERSONAL.XLSB`）
   — `src\PERSONAL\`のモジュール（CSV集約・事前精査・差分比較等の業務メインモジュール一式）

**`modAlarmOps.bas`は両ブックに別々に複製されている**（`src\Alarm_Master\modAlarmOps.bas`と
`src\PERSONAL\modAlarmOps.bas`は別ファイル・別内容。修正時はどちらのブック向けか必ず確認し、
両方に関わる修正であれば両方に反映する）。

これらのブックへのVBAモジュール反映を、「バックアップ → Import → コンパイル検証 →
失敗時ロールバック」の手順で安全に自動実行する。本番ファイルは直接編集せず、
必ず本ディレクトリ経由（PowerShell経由）で更新する。

## ディレクトリ構成

```
AlarmMasterAutomation\
├── CLAUDE.md                        ← このファイル
├── docs\
│   └── vba-alarm-tooling-SKILL.md   ← コーディング規約の正本（Read必須）
├── .github\                         ← GitHub Actions（2026-08-02追加）
│   ├── workflows\vba-lint.yml       ← push/PR時にsrc\配下の.bas/.clsを自動チェック
│   └── scripts\Check-VbaConventions.ps1
├── src\                             ← 現行.bas/.cls/.iniのマスターコピー
│   ├── (直下)                      ← Alarm_Master.xlsm向け（ThisWorkbook.cls・modSrcExport.bas等含む）
│   ├── Alarm_Master\                ← Alarm_Master.xlsm専用（modAlarmOps.bas等、直下と重複する場合あり）
│   └── PERSONAL\                    ← PERSONAL.XLSB専用（CSV_NZEd_NetAct.bas・modSrcExport.bas等）
├── bas_staging\                     ← 今回反映する差分だけを置く（対象ブックに応じてsrc\の該当フォルダからコピー。.bas/.cls両対応）
├── Update-AlarmMaster.ps1           ← 本体スクリプト（勝手に書き換えない）
├── Run_AlarmMaster_Update.bat       ← 手動実行用ランチャー
└── _backup\                         ← 自動生成。世代バックアップ＋ログ
    ├── update_log_*.txt
    └── export_log_AlarmMaster.txt / export_log_PERSONAL.txt  ← modSrcExport.basのログ（2026-08-02追加）
```

## VBAソース自動エクスポート（2026-08-02追加）

`src\`→ブックへの反映（本ワークフローの主目的）とは逆方向に、**両ブックとも保存時に`modSrcExport.bas`が全VBAモジュールを`src\`へ自動Exportする**仕組みが動いている。VBE上で直接編集された変更が`src\`に反映されないまま失われる事故を防ぐのが目的（自動commit/pushはしない。詳細はSKILL.md §5.7）。

このため、ユーザーがExcel上でブックを保存した後は、`git status`に**Claude Codeが関与していない差分**が現れることがある。これは異常ではないが、通常のsrc編集より一段慎重に扱うこと：
- 差分の内容を読み、意味のある変更か、VBE内部の表記正規化（SKILL.md §5.7の既知の癖）に過ぎないかを判断する。
- 台帳.cls（内部コンポーネント名`Sheet1`）のような、`src\`の慣習的ファイル名と実際のVBAコンポーネント名が食い違うケースでは重複ファイルが生成されうる（SKILL.md §5.7未解決課題）。安易に両方コミットせず、ユーザーに確認する。
- コミットするかどうかは都度ユーザーに確認する（無断でコミットしない）。

## 標準ワークフロー

1. ユーザーからの依頼を受けたら、まず `docs\vba-alarm-tooling-SKILL.md` と
   `src\` 内の対象モジュール（＋影響しそうな関連モジュール）を読む。
2. 修正方針・影響範囲・変更しないモジュールを箇条書きでユーザーに提示し、
   合意を得てから実装する（いきなりコードを書かない）。
3. 対象ブックを確認した上で、`src\`（Alarm_Master.xlsm向け）または`src\PERSONAL\`
   （PERSONAL.XLSB向け）内の該当ファイルを修正する。`modAlarmOps.bas`は両ブックに
   影響する修正なら`src\Alarm_Master\`と`src\PERSONAL\`の両方に反映する。
   同時に `bas_staging\` にも同じ内容を配置する（前回分が残っていれば削除・退避してから置く）。
4. 自己レビューを行い、以下を確認する：
   - `Option Explicit` の有無
   - `Const`/`Declare`/モジュールレベル`Dim`がモジュール先頭にあるか
   - 配列引数の`ByRef`明示
   - COMオブジェクトの明示解放（`Set obj = Nothing`）
   - 列参照が見出し名から取得されているか（ハードコード禁止）
   - Shift-JIS(CP932)保存・CRLF改行になっているか
5. 次のコマンドを実行する:
   ```
   powershell -NoProfile -ExecutionPolicy Bypass -File Update-AlarmMaster.ps1 `
       -TargetWorkbook "<対象ブックのフルパス>" -BasFolder "bas_staging"
   ```
   - `-RunAfterMacro` は**ユーザーが明示的に指定した場合のみ**付ける。
   - `-CompileTimeoutSec`（既定30秒）は、`Z_CompileCheck`等が応答しない場合に
     強制終了するまでの秒数。通常は指定不要（既定値で運用）。
   - `-RemoveModules "モジュール名,モジュール名"`（2026-07-31追加）: 対応する`.bas`
     無しで、指定したVBAコンポーネントを削除だけしたい場合に使う（誤って
     取り込んだモジュールの是正用）。バックアップ・コンパイル検証・
     失敗時ロールバックは通常のImportと同じ枠組みで適用される。
   - `-SkipCompileCheck`（2026-08-01追加）: `PERSONAL.XLSB`は別プロセスの
     Excelから開くと`Application.Run`によるマクロ実行が一切機能しない
     （空のマクロですら失敗する。原因未特定）ことを実機確認したため、
     `PERSONAL.XLSB`への反映では必ずこのスイッチを付ける。Import/保存の
     成否のみで成功判定するため、反映後のVBEでの手動確認が今まで以上に必須。
   - `bas_staging\`には`.bas`だけでなく`.cls`も置ける（2026-08-02追加）。
     `ThisWorkbook`等のドキュメントモジュールは`VBComponents.Remove`できない
     ため、通常のクラスモジュールとは異なりCodeModule差し替えで反映される
     （詳細はSKILL.md §5.7）。
6. 終了コード（`$LASTEXITCODE`）を確認する。
   - `0` = 成功。`_backup\update_log_*.txt` の末尾を要約して報告する。
     ただし**これはプロジェクト全体のコンパイル保証ではない**（後述・SKILL.md §5.6参照）。
   - `1` = 失敗。ロールバック済みである旨と、ログのHRESULT分類
     （`0x800A9C68`=コンパイルエラーの可能性、`0x800A03EC`=マクロ不明、
     タイムアウト=ブレークモードで無応答になった可能性）およびエラー内容をそのまま提示する。
     その上で、直前の変更をSKILL.mdのチェックリストに照らして自分で再レビュー・
     再修正し、再実行してよい。**この自己レビュー→再修正→再実行は連続最大3回まで**
     （2026-07-31合意。詳細はSKILL.md §5.6「VBA修正時の標準ループ」）。
     3回失敗したら止めてユーザーに判断を仰ぐ。
7. 成功時・失敗時いずれも、**最終的な全体コンパイル確認はユーザーによるVBEでの
   手動`Debug → Compile VBAProject`が必須**である旨を伝える（自動検証では代替できない
   ことを2026-07-31に実機検証済み。SKILL.md §5.6参照）。
8. 成功したら `docs\vba-alarm-tooling-SKILL.md` に仕様変更があれば同じセッションで
   更新を提案する（実装と規約集の乖離を残さない）。
9. Git管理している場合、コンパイル成功が確認できたバージョンのみコミットする
   （壊れた版をコミットしない）。この運用は`.githooks/pre-commit`（2026-08-02追加）
   でも機械的に補強されている：`.bas`/`.cls`をコミットしようとした際、対象ブック
   （`src/PERSONAL\`配下かどうかで判定）の直近`Update-AlarmMaster.ps1`実行ログ
   （対象ブックの実際の場所に出力される。プロジェクト直下の`_backup\`ではない）
   が「異常終了」ならコミットをブロックする。ログが見つからない場合はブロックせず
   警告のみ（誤検知回避）。有効化には`git config core.hooksPath .githooks`が必要
   （clone後1回のみ。このリポジトリでは既に設定済み）。意図的に回避する場合は
   `git commit --no-verify`（乱用しない）。

## 絶対に守ること（ガードレール）

- 本番ブックを、このスクリプトを介さずに直接VBEやCOM経由で書き換えない。
- 同一の失敗に対して**連続3回を超える自動リトライをしない**。3回失敗したら
  作業を止めてユーザーに判断を仰ぐ。
- `bas_staging\` 以外のフォルダにある `.bas`/`.cls` を確認なしに取り込まない。
- `_backup\` 配下のファイルを削除・上書きしない。
- `Update-AlarmMaster.ps1` の仕様変更は、必ず変更内容を提示してから行う。
- `modSrcExport.bas`によるブック保存時の自動Export差分（`git status`に現れる、
  Claude Codeが関与していない変更）を無断でcommitしない。必ず内容をレビューし、
  ユーザーに確認してからcommitする（詳細は前掲「VBAソース自動エクスポート」節、
  SKILL.md §5.7）。
- 生データシート（NZEd/NetAct）を直接改変するロジックを書かない。
- Alarm_Master.xlsm側の外部アプリ起動処理は `Shell.Application`/`WScript.Shell` を
  避け、Win32 API `ShellExecute` 直接呼び出しを踏襲する。
- MCP経由の `excel`（excel-mcp-server）は内部で `openpyxl.load_workbook()` を
  `keep_vba=True` 指定なしで使用しているため、`.xlsm` を開いて保存するとVBAプロジェクト
  （マクロ）が失われる。`Alarm_Master.xlsm` など既存の`.xlsm`に対しては**読み取り専用系
  ツール**（`read_data_from_excel`/`get_workbook_metadata`等）のみ使用可とし、
  書き込み系ツール（`write_data_to_excel`/`apply_formula`/`format_range`/
  `insert_rows`等、保存を伴うもの全般）は`.xlsm`に対して使用しない。CSV集約結果や
  事前精査テキストの整形など、補助的な別ファイル（`.xlsx`）を新規作成する用途に限る。
- `docs\vba-alarm-tooling-SKILL.md` に反する実装をした場合は、無断で規約側を
  緩めるのではなく、必ずユーザーに確認する。
- `PERSONAL.XLSB`はExcel起動時に常時自動読み込みされる個人用マクロブックのため、
  反映前に必ず`Get-Process EXCEL`等でExcelが起動していないか確認する
  （起動中に書き込みを試みるとファイルロック競合やユーザーの未保存作業への
  影響が起きるおそれがある。2026-07-31実機確認）。

## ユーザーへの報告フォーマット

```
■ 実行結果: 成功 / 失敗
■ 対象モジュール: <Import した.bas一覧>
■ コンパイル検証: OK / NG（NG時はエラー番号・メッセージ）
■ バックアップ: <世代パス>
■ SKILL.md/CLAUDE.md更新の要否: <あれば>
■ 次のアクション案: <あれば>
```
