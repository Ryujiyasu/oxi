// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! VBA's UI-language error catalog, independent of number/date formatting.

/// The UI language used by `Error()` and default `Err.Description` values.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum ErrorLanguage {
    #[default]
    English,
    Japanese,
}

/// Host-supplied defaults for errors raised by the VBA language.
///
/// `help_file` is the target host's installed help path. Leaving it empty is
/// suitable for a host without a help installation; no OS path is inferred.
#[derive(Debug, Clone, Default)]
pub struct ErrorEnvironment {
    pub language: ErrorLanguage,
    pub help_file: String,
}

impl ErrorEnvironment {
    pub fn description(&self, number: i64) -> &'static str {
        match self.language {
            ErrorLanguage::English => vba_error_description(number),
            ErrorLanguage::Japanese => japanese_error_description(number),
        }
    }

    pub(crate) fn automation_out_of_range(&self) -> String {
        let detail = match self.language {
            ErrorLanguage::English => "Out of present range. ",
            ErrorLanguage::Japanese => "範囲外です。 ",
        };
        format!("{}\n{detail}", self.description(440))
    }
}

pub(crate) fn vba_error_description(number: i64) -> &'static str {
    match number {
        0 => "",
        5 => "Invalid procedure call or argument",
        6 => "Overflow",
        7 => "Out of memory",
        9 => "Subscript out of range",
        10 => "This array is fixed or temporarily locked",
        11 => "Division by zero",
        13 => "Type mismatch",
        14 => "Out of string space",
        28 => "Out of stack space",
        35 => "Sub or Function not defined",
        48 => "Error in loading DLL",
        52 => "Bad file name or number",
        53 => "File not found",
        54 => "Bad file mode",
        55 => "File already open",
        58 => "File already exists",
        61 => "Disk full",
        62 => "Input past end of file",
        67 => "Too many files",
        68 => "Device unavailable",
        70 => "Permission denied",
        71 => "Disk not ready",
        75 => "Path/File access error",
        76 => "Path not found",
        91 => "Object variable or With block variable not set",
        92 => "For loop not initialized",
        93 => "Invalid pattern string",
        94 => "Invalid use of Null",
        424 => "Object required",
        429 => "ActiveX component can't create object",
        430 => "Class does not support Automation or does not support expected interface",
        432 => "File name or class name not found during Automation operation",
        438 => "Object doesn't support this property or method",
        440 => "Automation error",
        445 => "Object doesn't support this action",
        446 => "Object doesn't support named arguments",
        447 => "Object doesn't support current locale setting",
        448 => "Named argument not found",
        449 => "Argument not optional",
        450 => "Wrong number of arguments or invalid property assignment",
        451 => {
            "Property let procedure not defined and property get procedure did not return an object"
        }
        453 => "Specified DLL function not found",
        457 => "This key is already associated with an element of this collection",
        458 => "Variable uses an Automation type not supported in Visual Basic",
        // Measured: Err.Raise -5 and vbObjectError + 513 both describe
        // themselves so.
        number if number < 0 => "Automation error",
        _ => "Application-defined or object-defined error",
    }
}

fn japanese_error_description(number: i64) -> &'static str {
    match number {
        0 => "",
        3 => "Return に対応する GoSub がありません。",
        5 => "プロシージャの呼び出し、または引数が不正です。",
        6 => "オーバーフローしました。",
        7 => "メモリが不足しています｡",
        9 => "インデックスが有効範囲にありません。",
        10 => "この配列は固定されているか、または一時的にロックされています。",
        11 => "0 で除算しました。",
        13 => "型が一致しません。",
        14 => "文字列領域が不足しています。",
        16 => "式が複雑すぎます。",
        17 => "要求された操作は実行できません。",
        18 => "ユーザーによる割り込みが発生しました。",
        20 => "エラーが発生していないときに Resume を実行することはできません。",
        28 => "スタック領域が不足しています。",
        35 => "Sub または Function が定義されていません。",
        47 => "DLL のクライアント アプリケーションの数が多すぎます。",
        48 => "DLL 読み込み時のエラーです。",
        49 => "DLL が正しく呼び出せません。",
        51 => "内部エラーです。",
        52 => "ファイル名または番号が不正です。",
        53 => "ファイルが見つかりません。",
        54 => "ファイル モードが不正です。",
        55 => "ファイルは既に開かれています。",
        57 => "デバイス I/O エラーです。",
        58 => "既に同名のファイルが存在しています。",
        59 => "レコード長が一致しません。",
        61 => "ディスクの空き容量が不足しています。",
        62 => "ファイルにこれ以上データがありません。",
        63 => "レコード番号が不正です。",
        67 => "ファイルが多すぎます。",
        68 => "デバイスが準備されていません。",
        70 => "書き込みできません。",
        71 => "ディスクが準備されていません。",
        74 => "ディスク名は変更できません。",
        75 => "パス名が無効です。",
        76 => "パスが見つかりません。",
        91 => "オブジェクト変数または With ブロック変数が設定されていません。",
        92 => "For ループが初期化されていません。",
        93 => "パターン文字列が不正です。",
        94 => "Null の使い方が不正です。",
        96 => "サポートされているイベント受信最大数のイベントが既に発生しているので、オブジェクトのイベント シンクを実行できません。",
        97 => "オブジェクトが定義クラスのインスタンスではない場合、このオブジェクトに関するフレンド関数は呼び出せません。",
        98 => "プロパティまたはメソッドの呼び出しの場合には、引数または戻り値としてプライベート オブジェクトへの参照を含めることはできません。",
        321 => "不正なファイル形式です。",
        322 => "必要な一時ファイルを作成できません。",
        325 => "リソース ファイルの形式が不正です。",
        380 => "プロパティの値が不正です。",
        381 => "不正なプロパティ配列インデックスです。",
        382 => "実行時には値を設定できません。",
        383 => "値を設定できません。値の取得のみ可能なプロパティです。",
        385 => "プロパティ配列インデックスが必要です。",
        387 => "値を設定できません。",
        393 => "実行時には値を取得できません。",
        394 => "値を取得できません。値の設定のみ可能なプロパティです。",
        422 => "プロパティが見つかりません。",
        423 => "プロパティまたはメソッドが見つかりません。",
        424 => "オブジェクトが必要です。",
        429 => "ActiveX コンポーネントはオブジェクトを作成できません。",
        430 => "クラスはオートメーションまたは予測したインターフェースをサポートしていません。",
        432 => "オートメーションの操作中にファイル名またはクラス名を見つけられませんでした。",
        438 => "オブジェクトは、このプロパティまたはメソッドをサポートしていません。",
        440 => "オートメーション エラーです。",
        442 => "リモート プロセス用のタイプ ライブラリまたはオブジェクト ライブラリへの参照は失われました。参照設定を解除して [OK] を押してください。",
        443 => "オートメーション オブジェクトには既定値がありません。",
        445 => "オブジェクトはこの動作をサポートしていません。",
        446 => "オブジェクトは名前付き引数をサポートしていません。",
        447 => "オブジェクトは現在の国別情報の設定をサポートしていません。",
        448 => "名前付き引数が見つかりません。",
        449 => "引数は省略できません。",
        450 => "引数の数が一致していません。または不正なプロパティを指定しています。",
        451 => "Property Let プロシージャが定義されておらず、Property Get プロシージャからオブジェクトが返されませんでした。",
        452 => "序数が不正です。",
        453 => "関数は指定された DLL には定義されていません。",
        454 => "コード リソースが見つかりません。",
        455 => "コード リソースのロック エラーです。",
        457 => "このキーは既にこのコレクションの要素に割り当てられています。",
        458 => "Visual Basic でサポートされていないオートメーションが変数で使用されています。",
        459 => "オブジェクトまたはクラスがこのイベント セットをサポートしていません。",
        460 => "クリップボードのデータ形式が不正です。",
        461 => "メソッドまたはデータ メンバーが見つかりません。",
        462 => "リモート サーバーがないか、使用できる状態ではありません。",
        463 => "ローカル マシンにクラスが登録されていません。",
        481 => "ピクチャが不正です。",
        482 => "プリンター エラーです。",
        735 => "一時ファイルに保存できません。",
        744 => "検索文字列が見つかりませんでした。",
        746 => "置換後の文字列が長すぎます。",
        number if number < 0 => japanese_error_description(440),
        _ => "アプリケーション定義またはオブジェクト定義のエラーです。",
    }
}
