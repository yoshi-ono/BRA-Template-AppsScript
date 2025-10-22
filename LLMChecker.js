// Global constants

// --- CONFIGURATION FLAGS ---
const SHOW_ONLY_ERRORS = true; // true: エラー対象のみ表示, false: 全て表示

// --- GEMINI API SETUP ---
// ConfigシートからAPIキーとモデル名を取得する関数
function getConfigFromSheet() {
  try {
    const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
    if (!configSheet) {
      throw new Error('Configシートが見つかりません。');
    }
    
    const apiKey = configSheet.getRange('C6').getValue();
    const modelName = configSheet.getRange('C7').getValue();
    
    if (!apiKey || apiKey.toString().trim() === '') {
      throw new Error('ConfigシートのC6にAPIキーが入力されていません。');
    }
    
    if (!modelName || modelName.toString().trim() === '') {
      throw new Error('ConfigシートのC7にモデル名が入力されていません。');
    }
    
    return {
      apiKey: apiKey.toString().trim(),
      modelName: modelName.toString().trim()
    };
  } catch (e) {
    throw new Error(`設定の取得に失敗しました: ${e.message}`);
  }
}

// デフォルト値（Configシートが利用できない場合のフォールバック）
const DEFAULT_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
const DEFAULT_MODEL_NAME = 'gemini-2.5-flash';

const SYSTEM_PROMPT_CONTEXT = `
# あなたのタスク
あなたは、提供された「入力データ」を、指定された「評価対象エラー」に基づいて評価するAIアシスタントです。
評価結果を、必ず指定されたJSON形式で出力してください。

# 評価に必要な背景情報

## 1. 前提概念
現在の神経科学では、脳の様々な領域の解剖学的構造の理解が進んでいます。脳と似た認知行動を再現する神経回路モデルは、その計算機能が解剖学的構造と整合することで妥当性が高まります。
脳参照アーキテクチャ(BRA)駆動開発では、BRAという標準形式で計算モデルを記述し、脳型ソフトウェア開発を進めます。

## 2. BRAデータ
- **BRAデータ**: 標準化された脳型ソフトウェア記述形式。脳情報フロー（BIF）と仮説的コンポーネント図（HCD）から構成。
- **BIF (Brain Information Flow)**: 脳内の解剖学的構造を「Circuit」（ノード）と「connection」（リンク）で表現した有向グラフ。
- **HCD (Hypothetical Component Diagram)**: BIFの構造に整合するように機能を分解した仮説的コンポーネント図。
- **FRG (Function Realization Graph)**: 機能分解を行うための機能階層図。

## 3. 機能に関連する概念
- **Requirement**:
    - 定義: 機能ノードに対する要求機能。
    - 性質: 入出力信号の変換能力(Capability)と信号の意味付け(Output semantics)で規定。
    - 分解: 「役割分解」（特化したOutput semantics付与）と「体系的分解」（Capability分割）。
    - 記述例: 「[入力信号の意味]から[出力信号の意味]への変換を[Capability]により実現する」。例: 「生物刺激から恐怖応答への変換を条件付け学習により実現する」。
    - 記法: R.(Requirement名)
- **Capability**:
    - 定義: 機能ノード内部の信号変換能力。
    - 性質: 入力信号から出力信号への外形的な変換処理を定義。信号の意味付けとは無関係。
    - 記述例: 「[入力信号パターン]を[出力信号パターン]に変換する」。例: 「時系列信号を二値信号に変換する」。
    - 記法: C.(Capability名)
- **Output semantics**:
    - 定義: 外部観測者による信号パターンへの意味付け。
    - 性質: 出力信号に外部観測者から意味を付与。同一Uniform Circuitsの異なる階層間で意味づけの一貫性が必要。
    - 記述例: 「[信号パターン]は[意味]を表す」。例: 「高頻度発火は危険物の存在を表す」。
- **Mechanism**:
    - 定義: 機能ノード内のSubnodes処理の相互作用によるCapability実現の説明。
    - 性質: Subnodes処理の相互作用から得られる機能を自然言語で説明。個々のSubnode処理を超えた創発的機能を含む場合あり。Capabilityの実現を説明。階層の深掘りはしない。
    - 記述例: 「[Subnode1]と[Subnode2]の相互作用により[創発機能]が実現され、その結果[入力信号]が[出力信号]に変換される」。
- **Implementation**:
    - 定義: 機能ノード内における信号処理の形式的表現の列挙。
    - 性質: 疑似コードによる具体的な処理手順の列挙。内部の個別変換処理を形式的に記述。
    - 記法例: \`[ STR ] = U.STR( dmPFC )\`
- **Interface**:
    - 定義: 機能ノードの入力接続と出力信号の定義。
    - 性質: 接続される入力信号と出力信号を定義。HCD上の信号の流れを規定。
- **Uniform Circuits**:
    - 定義: BIF上で定義されたCircuitに基づく。一つの信号のOutput semanticsを対応付ける単位。
    - 記法: U.(Uniform-Circuit名)

## 4. 機能の評価に関連する概念
- **Requirement realization by interface**: 根拠に基づいた、Interfaceにおける入出力のOutput semanticsによるRequirementの実現可能性の説明。
- **Mechanism (記述要件として)**: Capabilityを実現するための、Implementationとその計算順序を含めた説明。

# あなたへの指示

以下の「入力データ」と「評価対象エラー」の内容をよく読み、評価を行ってください。
評価結果は、必ず下記のJSON形式で、そのJSON文字列のみを出力してください。

## 出力JSON形式
\`\`\`json
{
  "error_id": "評価したエラーコードのID (例: 1001)",
  "is_error_found": true,
  "reason": "エラーであると判断した場合、その具体的な理由。エラーでない場合は '問題なし' と記述。",
  "suggestion": "エラーであると判断した場合の改善案。エラーでない場合はnullまたは空文字列。"
}
\`\`\`
is_error_found は、エラーが実際に存在する場合に true、存在しない場合に false としてください。
reason と suggestion は、具体的かつ簡潔に記述してください。
重要事項
出力は、上記で指定されたJSON形式の文字列のみとしてください。
JSONの前後に説明文、コメント、マークダウンの json タグなど、他のテキストは一切含めないでください。
`;

const ERROR_DEFINITIONS = {
    "1001": {
        "target column": ["Requirements realization by interface"],
        "説明": "先行する実現事例や理論が示されていない",
        "具体例": ["「前時刻の位置情報と嗅覚情報を入力とする回路は、これらの入力から次時刻の位置情報を計算し出力できることは過去の文献でも指摘されており、実現可能である。」※文献情報がない。", "Mechanism カラムでの裏付け説明（計算手順など）が未記載、または『未記入』のまま。        "],
        "カラムの種類": "Requirements realization by interface",
        "カラムの説明": ["Interface カラム経由で Requirement をどのように満たすかを記述するカラム。"]
    },
    "1101": {
        "target column": ["Requirements", "Interface"],
        "説明": "Interface に記載された入力 Circuit が Requirements に記載された入力 Uniform Circuit と一致していない",
        "具体例": ["Requirements:『前時刻の位置情報と嗅覚情報を入力…』、Interface:『U.YYY(XXX1, XXX2, XXX3)』で Circuit 名不一致。"],
        "カラムの種類": "Requirements, Interface",
        "カラムの説明": ["Requirements: 機能ノードに対する要求機能を、Capability と Output semantics で規定するカラム。","Interface: 機能ノードの入力信号と出力信号を定義し、信号の流れを規定するカラム。"]
    },
    "1102": {
        "target column": ["Requirements", "Interface"],
        "説明": "Requirements に記載された出力 Uniform Circuit が Interface の出力 Uniform Circuit に含まれていない",
        "具体例": ["Requirements には出力 C があるのに Interface の出力リストに C が存在しない。"],
        "カラムの種類": "Requirements, Interface",
        "カラムの説明": ["Requirements: 機能ノードに対する要求機能を、Capability と Output semantics で規定するカラム。","Interface: 機能ノードの入力信号と出力信号を定義し、信号の流れを規定するカラム。"]
    },
    "1103": {
        "target column": ["Requirements realization by interface", "Output semantics"],
        "説明": "Requirements realization by interface に記述した入出力の Output semantics が当該 Uniform Circuit の説明と一致しない",
        "具体例": ["出力 Y の意味を『速度』と記載しているが、Uniform Circuit 側では『角速度』として定義されている。"],
        "カラムの種類": "Requirements realization by interface, Output semantics",
        "カラムの説明": ["Requirements realization by interface: Interface 経由で Requirement をどのように満たすかを記述するカラム。","Output semantics: 外部観測者による出力信号パターンへの意味付けを記述するカラム。"]
    },
    "1104": {
        "target column": ["Requirements realization by interface", "Requirements"],
        "説明": "Requirements realization by interface に記載された入出力が Requirements の記述と一致しない",
        "具体例": ["Requirements に無い追加入力が Requirements realization by interface にだけ登場している。"],
        "カラムの種類": "Requirements realization by interface, Requirements",
        "カラムの説明": ["Requirements realization by interface: Interface 経由で Requirement をどのように満たすかを記述するカラム。","Requirements: 機能ノードに対する要求機能を、Capability と Output semantics で規定するカラム。"]
    },
    "2001": {
        "target column": ["Mechanism"],
        "説明": "Input の Uniform Circuit が示されていない",
        "具体例": ["Mechanism カラムに『入力: (InputCircuit)』等の記載がなく、どの回路を入力にするか不明。"],
        "カラムの種類": "Mechanism",
        "カラムの説明": ["Mechanism: Subnodes の相互作用により Capability を実現する仕組みを自然言語で説明するカラム。"]
    },
    "2002": {
        "target column": ["Mechanism"],
        "説明": "Output の Uniform Circuit が示されていない",
        "具体例": ["Mechanism カラムに『出力: (OutputCircuit)』等の記載がなく、どの回路が出力か不明。"],
        "カラムの種類": "Mechanism",
        "カラムの説明": ["Mechanism: Subnodes の相互作用により Capability を実現する仕組みを自然言語で説明するカラム。"]
    },
    "2101": {
        "target column": ["Mechanism", "Capability"],
        "説明": "Mechanism の説明が Capability を実現していない",
        "具体例": ["Capability で『学習率の自動調整が可能』と宣言しているが、Mechanism で学習率更新の手順が説明されていない。"],
        "カラムの種類": "Mechanism, Capability",
        "カラムの説明": ["Mechanism: Subnodes の相互作用により Capability を実現する仕組みを説明するカラム。","Capability: 入力信号を出力信号へ外形的に変換する能力を記述するカラム。"]
    },
    "2102": {
        "target column": ["Mechanism", "Implementation"],
        "説明": "Mechanism の説明が Implementation に含まれる内容を網羅していない",
        "具体例": ["Implementation に具体的な前処理ステップがあるのに Mechanism でその記載が抜けている。"],
        "カラムの種類": "Mechanism, Implementation",
        "カラムの説明": ["Mechanism: Subnodes の相互作用により Capability を実現する仕組みを説明するカラム。","Implementation: 疑似コードなどで具体的な信号処理手順を列挙するカラム。"]
    },
    "2103": {
        "target column": ["Mechanism", "Implementation"],
        "説明": "Mechanism が Implementation の処理順序（計算的フロー）を含んでいない",
        "具体例": ["Implementation には『Step1 -> Step2 -> Step3』とあるが Mechanism に順序の言及がない。"],
        "カラムの種類": "Mechanism, Implementation",
        "カラムの説明": ["Mechanism: Subnodes の相互作用により Capability を実現する仕組みを説明するカラム。","Implementation: 疑似コードなどで具体的な信号処理手順を列挙するカラム。"]
    },
    "2104": {
        "target column": ["Mechanism", "Output semantics"],
        "説明": "Output semantics の内容が Mechanism に書かれている",
        "具体例": ["Mechanism に『この出力は角速度を表す』とあり、本来 Output semantics に記載すべき説明が混在している。"],
        "カラムの種類": "Mechanism, Output semantics",
        "カラムの説明": ["Mechanism: Subnodes の相互作用により Capability を実現する仕組みを説明するカラム。","Output semantics: 外部観測者による出力信号パターンへの意味付けを記述するカラム。"]
    },
    "2201": {
        "target column": ["Capability", "Output semantics"],
        "説明": "Output semantics の内容が Capability に書かれている",
        "具体例": ["Capability に『出力 V は…』と出力の意味を詳述しており、Output semantics カラムが空欄。"],
        "カラムの種類": "Capability, Output semantics",
        "カラムの説明": ["Capability: 入力信号を出力信号へ外形的に変換する能力を記述するカラム。","Output semantics: 外部観測者による出力信号パターンへの意味付けを記述するカラム。"]
    },
    "2202": {
        "target column": ["Requirements", "Output semantics"],
        "説明": "Requirements に Output semantics が書かれていない",
        "具体例": ["Requirements で各出力を定義しているが、その意味や単位が Output semantics にも Requirements にも未記載。"],
        "カラムの種類": "Requirements, Output semantics",
        "カラムの説明": ["Requirements: 機能ノードに対する要求機能を、Capability と Output semantics で規定するカラム。","Output semantics: 外部観測者による出力信号パターンへの意味付けを記述するカラム。"]
    },
    "2203": {
        "target column": ["Output semantics", "Output semantics (Uniform Circuit)"],
        "説明": "Output semantics と Output semantics (Uniform Circuit) の記述が整合していない",
        "具体例": ["Output semantics に『速度』、Uniform Circuit では『角速度』と定義されている。"],
        "カラムの種類": "Output semantics, Output semantics (Uniform Circuit)",
        "カラムの説明": ["Output semantics: 外部観測者による出力信号パターンへの意味付けを記述するカラム。","Output semantics (Uniform Circuit): Uniform Circuit 単位での出力信号の意味を記述し、階層間で一貫性を確保するカラム。"]
    },
    "2204": {
        "target column": ["Requirements", "Capability", "Output semantics"],
        "説明": "Capability に Output semantics を入れても Requirements を満たす説明になっていない",
        "具体例": ["Capability に詳細な出力意味を記載しているが、Requirements の機能要求と直接対応付ける記述が無い。"],
        "カラムの種類": "Requirements, Capability, Output semantics",
        "カラムの説明": ["Requirements: 機能ノードに対する要求機能を、Capability と Output semantics で規定するカラム。","Capability: 入力信号を出力信号へ外形的に変換する能力を記述するカラム。","Output semantics: 外部観測者による出力信号パターンへの意味付けを記述するカラム。"]
    },
    "2205": {
        "target column": ["Mechanism"],
        "説明": "Mechanism カラム内で他行と重複・矛盾する説明が含まれている",
        "具体例": ["同一カラム内で『積分器』と『微分器』の両方を同時に実装すると記述しているが、前後の説明が整合していない。"],
        "カラムの種類": "Mechanism",
        "カラムの説明": ["Mechanism: Subnodes の相互作用により Capability を実現する仕組みを自然言語で説明するカラム。"]
    },
    "2206": {
        "target column": ["Requirements", "Capability"],
        "説明": "Capability と Requirements が整合していない",
        "具体例": ["Requirements で処理速度 1 ms 未満と規定されているのに、Capability では 10 ms と記載している。"],
        "カラムの種類": "Requirements, Capability",
        "カラムの説明": ["Requirements: 機能ノードに対する要求機能を、Capability と Output semantics で規定するカラム。","Capability: 入力信号を出力信号へ外形的に変換する能力を記述するカラム。"]
    }
};

const ERROR_ID_LIST = ['1001', '1101', '1102', '1103', '1104',
                     '2001', '2002',
                     '2101', '2102', '2103', '2104',
                     '2201', '2202', '2204', '2206'];


const INPUT_COLUMN_MAPPING = {
  "Requirements realization by interface": 59,
  "Requirements": 61,
  "Interface": 51,
  "Output semantics": 63,
  "Mechanism": 53,
  "Capability": 56,
  "Implementation": 52,
  "Output semantics (Uniform Circuit)": 27
};

const OUTPUT_COLUMN_MAPPING = {
  "1001": 60,
  "1101": 58,
  "1102": 58,
  "1103": 58,
  "1104": 58,
  "2001": 54,
  "2002": 54,
  "2101": 54,
  "2102": 54,
  "2103": 54,
  "2104": 54,
  "2201": 54,
  "2202": 62,
  "2203": 64,
  "2204": 54,
  "2205": 57,
  "2206": 62
};


/**
 * onOpen ; スプレッドシートが開かれたときにカスタムメニューを追加します。
 * 削除以降
 */

/**
 * runFRGChecksOnSelectedRowWithUI ; UI経由でメイン処理を呼び出すためのラッパー関数
 * 削除以降
 */


/**
 * 選択された行のFRGデータをチェックし、結果をスプレッドシートに書き込みます。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 対象のシート。
 * @param {number} currentRow 対象の行番号。
 * @param {Object} config 設定オブジェクト（apiKey, modelName）。
 */
function runFRGChecksOnSelectedRow(sheet, currentRow, config) {
  const ui = SpreadsheetApp.getUi();
  let errorsEncountered = 0;
  
  // 結果を一時的に蓄積するオブジェクト（カラム番号をキーとして使用）
  const resultsToUpdate = {};
  
  // 処理開始時にA列セルを赤背景に変更
  const aColumnCell = sheet.getRange(currentRow, 1);
  const originalBackground = aColumnCell.getBackground();
  aColumnCell.setBackground('#ff0000'); // 赤背景

  try {
    for (let i = 0; i < ERROR_ID_LIST.length; i++) {
    const errorId = ERROR_ID_LIST[i];
    const errorDefinition = ERROR_DEFINITIONS[errorId];

    if (!errorDefinition) {
      Logger.log(`行${currentRow}: エラーID '${errorId}' の定義が見つかりません。スキップします。`);
      errorsEncountered++;
      continue;
    }

    SpreadsheetApp.getActiveSpreadsheet().toast(`行${currentRow}: エラーID ${errorId} を処理中... (${i + 1}/${ERROR_ID_LIST.length})`, `FRGチェック中`, 10);
    Logger.log(`Processing error_id: ${errorId} for row ${currentRow}`);

    // 1. 入力データ文字列の構築
    let inputDataString = "#入力\n";
    const targetColumns = errorDefinition["target column"];
    if (!targetColumns || targetColumns.length === 0) {
        Logger.log(`行${currentRow}: エラーID '${errorId}' に 'target column' が定義されていません。スキップします。`);
        errorsEncountered++;
        const errorResult = {
            error_id: errorId,
            is_error_found: true,
            reason: `スクリプト設定エラー: エラーID '${errorId}' に 'target column' が定義されていません。`,
            suggestion: "INPUT_COLUMN_MAPPING と ERROR_DEFINITIONS を確認してください。"
        };
        addResultToBatch(resultsToUpdate, errorId, errorResult);
        continue;
    }

    for (const colName of targetColumns) {
      const colNum = INPUT_COLUMN_MAPPING[colName];
      if (!colNum) {
        Logger.log(`行${currentRow}: 入力カラムマッピングにカラム名 '${colName}' (エラーID: ${errorId}) が見つかりません。スキップします。`);
        errorsEncountered++;
        const errorResult = {
            error_id: errorId,
            is_error_found: true,
            reason: `スクリプト設定エラー: 入力カラム '${colName}' のマッピングが見つかりません。`,
            suggestion: "INPUT_COLUMN_MAPPING を確認してください。"
        };
        addResultToBatch(resultsToUpdate, errorId, errorResult);
        inputDataString = null; // APIコールをスキップするフラグとして使用
        break; 
      }
      const cellValue = sheet.getRange(currentRow, colNum).getValue();
      inputDataString += `種類: ${colName}\n内容: ${String(cellValue)}\n\n`;
    }

    if (inputDataString === null) { // inputDataString構築中にエラーがあった場合
        continue; // 次のエラーIDへ
    }


    // 2. エラー詳細プロンプト文字列の構築
    let errorDetailsPromptString = `
エラーID: ${errorId}
説明: ${errorDefinition['説明'] || 'N/A'}
評価対象カラムの種類: ${errorDefinition['カラムの種類'] || 'N/A'}
具体例: ${(errorDefinition['具体例'] || []).join('\n') || 'N/A'}

あなたのタスクは、上記の「入力データ」が、「評価対象エラー」に該当するかどうかを判断し、システムプロンプトで指示されたJSON形式で結果を返すことです。
`;

    // 3. ユーザープロンプト全体の構築
    const userPrompt = `
# タスク概要
- あなたのタスクは、以下の「入力データ」を、システムプロンプトで提供された「エラーコードリスト」に基づいて網羅的にチェックすることです。 (このプロンプトでは単一のエラーコードを扱います)
- 指示された手順に従い、「最終的な出力形式 (JSON)」に従ってJSONデータのみを出力してください。
- 論理的かつ体系的に評価を進めてください。

# 入力データ
${inputDataString}
# 調査対象error
${errorDetailsPromptString}
`;

    // 4. Gemini API呼び出し
    let resultJsonStr;
    try {
      resultJsonStr = callGeminiAPI(SYSTEM_PROMPT_CONTEXT, userPrompt, config);
    } catch (e) {
      Logger.log(`行${currentRow}: Gemini API呼び出し中にエラー (エラーID: ${errorId}): ${e.toString()}\nStack: ${e.stack}`);
      errorsEncountered++;
      const errorResult = {
          error_id: errorId,
          is_error_found: true, // APIエラーなので問題ありとみなす
          reason: `API呼び出しエラー: ${e.message}`,
          suggestion: "APIキー、エンドポイント、ネットワーク接続を確認してください。"
      };
      addResultToBatch(resultsToUpdate, errorId, errorResult);
      continue; // 次のエラーIDへ
    }

    // 5. 結果のパースとバッチに追加
    if (resultJsonStr) {
      try {
        const resultObj = JSON.parse(resultJsonStr);
        if (typeof resultObj.error_id === 'undefined') { // LLMがerror_idを返さなかった場合
            resultObj.error_id = errorId;
        }
        addResultToBatch(resultsToUpdate, errorId, resultObj);
      } catch (e) {
        Logger.log(`行${currentRow}: JSONパースエラー (エラーID: ${errorId}): ${e.toString()}. API応答: ${resultJsonStr}`);
        errorsEncountered++;
        const errorResult = {
            error_id: errorId,
            is_error_found: true, // パースエラーなので問題ありとみなす
            reason: `API応答のJSONパースエラー: ${e.message}`,
            suggestion: "APIの応答形式を確認してください。",
            raw_output: resultJsonStr // 生の応答も記録しておくとデバッグに役立つ
        };
        addResultToBatch(resultsToUpdate, errorId, errorResult);
      }
    } else {
      Logger.log(`行${currentRow}: Gemini APIから空の応答 (エラーID: ${errorId})`);
      errorsEncountered++;
      const errorResult = {
          error_id: errorId,
          is_error_found: true, // 空の応答は問題あり
          reason: "APIから空の応答がありました。",
          suggestion: "APIの状態やプロンプトを確認してください。"
      };
      addResultToBatch(resultsToUpdate, errorId, errorResult);
    }

    // APIのレート制限を避けるために少し待機
    Utilities.sleep(1500); // 1.5秒待機 (必要に応じて調整)
  }

    // 6. 全ての処理完了後、一括でスプレッドシートに書き込み
    Logger.log("全てのエラーID処理が完了しました。結果を一括でスプレッドシートに書き込みます。");
    batchUpdateResults(sheet, currentRow, resultsToUpdate);

    if (errorsEncountered > 0) {
      Logger.log(`行${currentRow}: ${errorsEncountered} 件のエラー処理中に問題が発生しました。詳細はログを確認してください。`);
      // ui.alert("注意", `${errorsEncountered} 件のエラー処理中に問題が発生しました。詳細はログを確認してください。`, ui.ButtonSet.OK);
    }
  } catch (e) {
    Logger.log(`行${currentRow}: FRGチェック処理中にエラーが発生しました: ${e.toString()}\nStack: ${e.stack}`);
    throw e; // エラーを再スローして上位のエラーハンドリングに委ねる
  } finally {
    // 処理完了時（成功・失敗問わず）にA列セルの背景を元に戻す
    aColumnCell.setBackground(originalBackground);
  }
}

/**
 * 結果をバッチ更新用のオブジェクトに追加します。
 * @param {Object} resultsToUpdate バッチ更新用の結果オブジェクト。
 * @param {string} errorId エラーID。
 * @param {Object} resultObj 結果オブジェクト。
 */
function addResultToBatch(resultsToUpdate, errorId, resultObj) {
  const outputColumn = OUTPUT_COLUMN_MAPPING[errorId];
  if (!outputColumn) {
    Logger.log(`エラーID '${errorId}' の出力カラムマッピングが見つかりません。結果をバッチに追加できません。`);
    return;
  }

  if (!resultsToUpdate[outputColumn]) {
    resultsToUpdate[outputColumn] = [];
  }
  
  resultsToUpdate[outputColumn].push(resultObj);
  Logger.log(`エラーID ${errorId} の結果をカラム ${outputColumn} のバッチに追加しました。`);
}

/**
 * 蓄積された結果を一括でスプレッドシートに書き込みます。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 対象のシート。
 * @param {number} row 書き込む行番号。
 * @param {Object} resultsToUpdate バッチ更新用の結果オブジェクト（カラム番号をキーとする）。
 */
function batchUpdateResults(sheet, row, resultsToUpdate) {
  Logger.log(`バッチ更新開始: 行 ${row} に ${Object.keys(resultsToUpdate).length} 個のカラムを更新します。`);
  
  for (const [columnNum, resultsArray] of Object.entries(resultsToUpdate)) {
    const colNum = parseInt(columnNum);
    if (!colNum || resultsArray.length === 0) {
      Logger.log(`カラム ${columnNum} の結果が空のためスキップします。`);
      continue;
    }

    try {
      const cell = sheet.getRange(row, colNum);
      
      let displayResults = resultsArray;
      let statusPrefix = "";
      
      if (SHOW_ONLY_ERRORS) {
        // エラー対象のみを表示
        displayResults = resultsArray.filter(result => result.is_error_found === true);
        if (displayResults.length === 0) {
          statusPrefix = "OK:";
        } else {
          statusPrefix = "NG:";
        }
      } else {
        // 全て表示
        const allErrorsFalse = resultsArray.every(result => result.is_error_found === false);
        statusPrefix = allErrorsFalse ? "OK:" : "NG:";
      }
      
      // JSON文字列を生成
      let jsonStringToWrite = JSON.stringify(displayResults, null, 2);
      
      // ステータスプレフィックスを冒頭に追加
      if (statusPrefix) {
        jsonStringToWrite = statusPrefix + "\n" + jsonStringToWrite;
      }
      
      // 上書き更新
      cell.setValue(jsonStringToWrite);

      Logger.log(`カラム ${colNum} に ${displayResults.length} 件の結果を上書きしました。ステータス: ${statusPrefix}`);

    } catch (e) {
      Logger.log(`カラム ${colNum} への書き込み中にエラーが発生しました: ${e.toString()}`);
    }
  }
  
  Logger.log("バッチ更新が完了しました。");
}

/**
 * Gemini APIを呼び出します。
 * @param {string} systemInstruction システムプロンプト。
 * @param {string} userPromptContent ユーザープロンプト。
 * @param {Object} config 設定オブジェクト（apiKey, modelName）。
 * @return {string|null} APIからのJSON応答文字列、またはエラーの場合はnull。
 */
function callGeminiAPI(systemInstruction, userPromptContent, config) {
  // 動的にAPIエンドポイントを構築
  const apiEndpoint = `https://generativelanguage.googleapis.com/v1beta/models/${config.modelName}:generateContent?key=${config.apiKey}`;
  const payload = {
    "contents": [{
      "role": "user", // system instructionは最上位のmodelオブジェクトに渡すため、ここではuserのみ
      "parts": [{ "text": userPromptContent }]
    }],
    "systemInstruction": { // system_instructionはここに設定
       "parts":[{"text": systemInstruction}]
    },
    "generationConfig": {
      "temperature": 0.0,
      "responseMimeType": "application/json" // JSON形式での出力を要求
    },
    "safetySettings": [ // Pythonコードに合わせてSAFETYをBLOCK_NONEに設定
      {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_NONE"},
      {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_NONE"},
      {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_NONE"},
      {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_NONE"},
    ]
  };

  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true // これでエラー時にもレスポンスボディを取得できる
  };

  Logger.log("Gemini API Request Payload: " + JSON.stringify(payload).substring(0, 500) + "..."); // Log first 500 chars of payload

  const response = UrlFetchApp.fetch(apiEndpoint, options);
  const responseCode = response.getResponseCode();
  const responseBody = response.getContentText();

  Logger.log("Gemini API Response Code: " + responseCode);
  Logger.log("Gemini API Response Body: " + responseBody.substring(0, 1000) + "..."); // Log first 1000 chars of response

  if (responseCode === 200) {
    try {
        // responseMimeType: "application/json" を指定しているので、
        // responseBody が直接JSON文字列のはず。
        // Python版の response.text に相当。
        const parsedResponse = JSON.parse(responseBody);
        if (parsedResponse.candidates && parsedResponse.candidates.length > 0 &&
            parsedResponse.candidates[0].content && parsedResponse.candidates[0].content.parts &&
            parsedResponse.candidates[0].content.parts.length > 0 && parsedResponse.candidates[0].content.parts[0].text) {
          return parsedResponse.candidates[0].content.parts[0].text.trim();
        } else if (parsedResponse.promptFeedback && parsedResponse.promptFeedback.blockReason) {
           Logger.log(`API Blocked Response: ${JSON.stringify(parsedResponse.promptFeedback)}`);
           throw new Error(`API request blocked. Reason: ${parsedResponse.promptFeedback.blockReason}`);
        }
        else {
          Logger.log(`API Success but unexpected response structure: ${responseBody}`);
          throw new Error("API returned 200 but the response structure was not as expected.");
        }
    } catch (e) {
        Logger.log(`Error parsing successful API response or unexpected structure: ${e.toString()}. Response body: ${responseBody}`);
        // もし responseMimeType が効かず、マークダウンでラップされたJSONが返ってきた場合のフォールバック
        const jsonMatch = responseBody.match(/```json\s*(\{[\s\S]*?\})\s*```/);
        if (jsonMatch && jsonMatch[1]) {
            Logger.log("Extracted JSON from markdown block.");
            return jsonMatch[1].trim();
        }
        throw new Error(`Failed to parse JSON from API response: ${e.message}. Original response was logged.`);
    }
  } else {
    Logger.log(`API Error. Response Code: ${responseCode}, Body: ${responseBody}`);
    let errorMessage = `Gemini API Error (Code: ${responseCode}).`;
    try {
      const errorJson = JSON.parse(responseBody);
      if (errorJson.error && errorJson.error.message) {
        errorMessage += ` Message: ${errorJson.error.message}`;
      } else {
        errorMessage += ` Body: ${responseBody.substring(0, 200)}...`;
      }
    } catch (e) {
      errorMessage += ` Raw Body: ${responseBody.substring(0, 200)}...`;
    }
    throw new Error(errorMessage);
  }
}

