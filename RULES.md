# RULES.md (Codex 運用ルール)

このリポジトリは「Excel の関数や加工処理を Python に置き換える」用途で運用する。
Codex は必ず本ファイルのルールに従うこと。
本ファイルは仕様・契約・前提条件そのものであり、要約や解釈を行ってはならない。

------------------------------------------------------------
1. 絶対ルール(勝手に変えない)
------------------------------------------------------------

(1) 既存ファイルについては、
    「仕様変更が決まり、
     その仕様に基づいて修正してください」
    と明示的に指示された部分以外は、
    一切変更してはならない。

    ・指示されていない行
    ・指示されていない関数
    ・指示されていないブロック
    ・その結果として影響が及ぶ周辺コード

    これらはすべて変更禁止とする。

(2) それ以外の部分は、一切いじらず、すべて残す。

    ・コメントは、一切いじらず、すべて残す
    ・インデントは、一切いじらず、すべて残す
    ・文字列の改行位置や空白も、指示がない限り勝手に整形しない。

(3) 絵文字・機種依存文字は禁止。
    使用する文字は ASCII と第一水準・第二水準漢字の範囲に限定する。
    (記号は ◇□■◆ などは使用してよい)

(4) 表形式で説明しない。
    (図は必要なら可)

------------------------------------------------------------
1-2. File Modification Policy (Editable / Reference Files)
------------------------------------------------------------

### Editable Files
仕様で指示したPythonファイル。

### Reference-Only Files (DO NOT MODIFY)

The following files are provided for reference only.
They must not be edited, reformatted, renamed, or deleted.

This prohibition includes even trivial changes such as
whitespace, indentation, comments, import order, or encoding.

- src/csv_to_tsv_h_mm_ss.py
- src/manhour_remove_uninput_rows.py
- src/sort_manhour_by_staff_code.py
- src/convert_yyyy_mm_dd.py
- src/make_staff_code_range.py
- src/make_sheet6_from_sheet4.py
- src/make_sheet789_from_sheet4.py
- src/make_unique_staff_code_list.py
- src/convert_salary_horizontal_to_vertical.py

Any change to files not explicitly listed under "Editable Files"
is strictly prohibited.

------------------------------------------------------------
1-3. Input / Expected Files Policy
------------------------------------------------------------

### Input Files (Read-Only)

The following files are input data only.
They must not be modified, renamed, moved, or deleted.

- input/Raw_Data.tsv
- input/Project_List_Formula.tsv
- input/支給・控除等一覧表_給与_2025年09月19日支給20251113.csv
- input/With_Salary_Formula.tsv

### Expected Output Files (Verification Only)

The following files represent expected correct results.
They are used only for comparison and verification.

- expected/支給・控除等一覧表_給与_2025年09月19日支給20251113_vertical.tsv
- expected/With_Salary.tsv

------------------------------------------------------------
1-4. Embedded Logic Policy (Reference Script Copying)
------------------------------------------------------------

目的:
- Reference-Only として指定された Python ファイルの処理内容を、
  Editable Files に「同一挙動」として内蔵することを許可する。

前提:
- Reference-Only Files 自体は、一切変更しない。

ルール:
(1) 内蔵とは「ロジックを複製して書き写す」ことを意味し、
    import して呼び出すことを意味しない。

(2) 内蔵された処理は、次の点について
    Reference-Only ファイルと完全一致していなければならない。

    ・入力の受け取り方法
    ・出力ファイル名および出力先
    ・データ変換内容
    ・例外発生時のメッセージ内容
    ・エラーファイルの生成規則
    ・終了条件

(3) csv_to_tsv_h_mm_ss.py の処理を内蔵する場合、
    csv_to_tsv_h_mm_ss.py の実装を正とし、
    RULES.md の一般規則よりも優先される。

(4) 内蔵のために変更してよい範囲は、
    依頼文で「仕様変更が決まり、その仕様に基づいて修正してください」
    と明示されたブロックのみとする。

(5) 内蔵によって他の既存処理への影響が発生する場合、
    Codex は作業を開始せず、必ず質問すること。

------------------------------------------------------------
2. 変数名と型(必須)
------------------------------------------------------------

(1) 変数には必ず型宣言を書くこと。
    Python でも型ヒントを必須とする。

    例:
    pszInputFileFullPath: str
    iRowCount: int
    objDataFrame: DataFrame

(2) 命名規則(接頭辞)

    ・short型: s
    ・int型: i
    ・unsigned int型: ui
    ・char型: c
    ・unsigned char型: uc
    ・long型: l
    ・BOOL型: b
    ・float型: f
    ・double型: d
    ・文字列(str): psz
    ・構造体(オブジェクト): obj
    ・構造体(ポインター): p

(3) 略しすぎ禁止

    ・意味が分かる英単語を省略しないこと。
    ・例:
      pszTf → NG
      pszTextFile → OK

    ・x, y, z を iX, iY, iZ のように1文字にせず、
      iCoordinateX, iCoordinateY, iCoordinateZ のように命名すること。

    構造体内の変数名も、同じ命名規則を適用する。

------------------------------------------------------------
3. コメント
------------------------------------------------------------

(1) コメントは、日本語で詳しく書く。

------------------------------------------------------------
(以下、bk01 / bk02 / 現行 RULES.md の内容をすべて包含済み)
