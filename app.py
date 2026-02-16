import streamlit as st
import pandas as pd
import google.generativeai as genai
import io
import time
from PIL import Image

# ==========================================
# ページ設定
# ==========================================
st.set_page_config(page_title="ICS年末調整データ作成ツール", layout="wide")

st.title("📄 ICS年末調整データ作成 AIツール (PDF対応・一括処理・データ統合版)")
st.markdown("""
複数の Excel(全シート)、CSV、**給与明細の画像・PDF** をまとめてアップロードできます。
**【新機能】**
1. Excelの**全てのシート**を読み取ります。
2. 複数ファイルに分かれた同一人物のデータを**自動で統合**します。
3. **「追加解析」**が可能になりました。ファイルを分けて投入してもデータが蓄積されます。
""")

# ==========================================
# セッション状態の初期化（データを記憶するため）
# ==========================================
# 読み取った全データを保持する場所を作成
if 'master_dataframe' not in st.session_state:
    st.session_state.master_dataframe = pd.DataFrame()

# サイドバー設定
st.sidebar.title("設定・操作")
api_key = st.sidebar.text_input("Google API Keyを入力", type="password")

# データリセットボタン
if st.sidebar.button("🗑️ 読み取りデータを全消去"):
    st.session_state.master_dataframe = pd.DataFrame()
    st.success("データをリセットしました。")

if api_key:
    genai.configure(api_key=api_key)

# ==========================================
# ファイルアップロード
# ==========================================
uploaded_files = st.file_uploader(
    "ファイルをまとめてドラッグ＆ドロップしてください", 
    type=["xlsx", "xls", "csv", "png", "jpg", "jpeg", "pdf"],
    accept_multiple_files=True
)

# ICSのヘッダー定義（データ処理用）
ICS_COLUMNS = [
    "個人コード", "区分コード", "氏名（姓）", "氏名（名）", "氏名フリガナ（姓）", "氏名フリガナ（名）",
    "性別", "入社年月日", "退職年月日", "郵便番号", "住所1", "住所2",
    "1月支給", "2月支給", "3月支給", "4月支給", "5月支給", "6月支給",
    "7月支給", "8月支給", "9月支給", "10月支給", "11月支給", "12月支給",
    "1月社保", "2月社保", "3月社保", "4月社保", "5月社保", "6月社保",
    "7月社保", "8月社保", "9月社保", "10月社保", "11月社保", "12月社保",
    "1月税額", "2月税額", "3月税額", "4月税額", "5月税額", "6月税額",
    "7月税額", "8月税額", "9月税額", "10月税額", "11月税額", "12月税額",
    "1月賞与", "2月賞与", "3月賞与", "4月賞与", "5月賞与", "6月賞与",
    "7月賞与", "8月賞与", "9月賞与", "10月賞与", "11月賞与", "12月賞与"
]

def process_single_content(content, source_name, file_type="text"):
    """AI解析処理（リトライ機能付き）"""
    model = genai.GenerativeModel('gemini-2.5-flash') 
    
    prompt_text = """

あなたは給与計算のプロフェッショナルです。

提供されたデータから給与情報を抽出し、ICS年末調整システム用のCSVデータを作成してください。



【最重要：出力ルール】

1. **ヘッダー行（項目名）は絶対に出力しないでください。** データの中身（数値・文字）の行のみを出力してください。

2. 以下の列順序（カンマ区切り）を厳守してください。

列構成: 個人コード,区分コード,氏名（姓）,氏名（名）,氏名フリガナ（姓）,氏名フリガナ（名）,性別,入社年月日,退職年月日,郵便番号,住所1,住所2,1月支給,2月支給額,3月支給額,4月支給額,5月支給額,6月支給額,7月支給額,8月支給額,9月支給額,10月支給額,11月支給額,12月支給額,1月社会保険料,2月社会保険料,3月社会保険料,4月社会保険料,5月社会保険料,6月社会保険料,7月社会保険料,8月社会保険料,9月社会保険料,10月社会保険料,11月社会保険料,12月社会保険料,1月所得税,2月所得税,3月所得税,4月所得税,5月所得税,6月所得税,7月所得税,8月所得税,9月所得税,10月所得税,11月所得税,12月所得税,1月賞与,2月賞与,3月賞与,4月賞与,5月賞与,6月賞与,7月賞与,8月賞与,9月賞与,10月賞与,11月賞与,12月賞与



3. 金額の「円」やカンマは削除し、半角数字のみにする。

4. 空白の項目はカンマだけ打つ（例: ,,,）。列ズレは許されません。

5. Markdownタグや挨拶は不要。CSVテキストデータのみを返してください。


添付されたファイルから給与データを抽出し、ICS年末調整システムの入力形式に合わせて出力してください。一個入力がずれると全てずれてしまうことを加味し、正確に情報を読み取り、正確にファイルを作成しなさい。



【出力ルール】

1. 形式はcsvとし、ヘッダー(項目名)を必ず含めること。

2. 行の構成：それぞれ個人ごとに行を改行する。

3. データがない項目は、項目名だけ書き、数字の場所は空欄(カンマのみ）にすること。

4. 数字にカンマや「円」を含めないこと。

5. 出力はcsvテキストのみ。余計な説明は一切不要

"""

    max_retries = 3
    for attempt in range(max_retries):
        try:
            if file_type == "pdf":
                pdf_data = {'mime_type': 'application/pdf', 'data': content}
                response = model.generate_content([prompt_text, pdf_data])
            elif file_type == "image":
                image = Image.open(io.BytesIO(content))
                response = model.generate_content([prompt_text, image])
            else:
                response = model.generate_content(prompt_text + f"\n\n【データソース: {source_name}】\n{content}")

            raw_text = response.text.replace("```csv", "").replace("```", "").strip()
            return raw_text

        except Exception as e:
            error_msg = str(e)
            if "429" in error_msg or "quota" in error_msg.lower() or "resource" in error_msg.lower():
                if attempt < max_retries - 1:
                    time.sleep(10) # 待機
                    continue
                else:
                    return f"ERROR: API制限により解析失敗 ({source_name})"
            else:
                return f"ERROR: {source_name} の解析に失敗: {e}"

# ==========================================
# 実行ボタン
# ==========================================
if uploaded_files and api_key:
    # ボタンのラベルを「追加解析」のニュアンスに変更
    if st.button(f"選択した {len(uploaded_files)} 件を追加解析・統合する", type="primary"):
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        new_data_rows = [] # 今回解析した分のリスト
        error_logs = []
        
        # 処理対象リストを作成（Excelのシート分解を含む）
        processing_targets = []
        
        # --- 1. ファイルの前処理（Excelシート分解） ---
        for file in uploaded_files:
            try:
                if file.name.endswith(('.xlsx', '.xls')):
                    # ★Excelの全シートを読み込む処理
                    excel_file = pd.ExcelFile(file)
                    for sheet_name in excel_file.sheet_names:
                        df = pd.read_excel(file, sheet_name=sheet_name)
                        csv_text = df.to_csv(index=False)
                        processing_targets.append({
                            "content": csv_text,
                            "name": f"{file.name} (Sheet: {sheet_name})",
                            "type": "text"
                        })
                elif file.name.endswith('.pdf'):
                    processing_targets.append({
                        "content": file.read(),
                        "name": file.name,
                        "type": "pdf"
                    })
                elif file.name.endswith(('.png', '.jpg', '.jpeg')):
                    processing_targets.append({
                        "content": file.read(),
                        "name": file.name,
                        "type": "image"
                    })
                elif file.name.endswith('.csv'):
                    df = pd.read_csv(file)
                    processing_targets.append({
                        "content": df.to_csv(index=False),
                        "name": file.name,
                        "type": "text"
                    })
            except Exception as e:
                error_logs.append(f"ファイル読み込みエラー: {file.name} - {e}")

        total_targets = len(processing_targets)

        # --- 2. AI解析ループ ---
        for i, target in enumerate(processing_targets):
            status_text.text(f"解析中 ({i+1}/{total_targets}): {target['name']} ...")
            
            result_text = process_single_content(target['content'], target['name'], target['type'])
            
            if result_text.startswith("ERROR"):
                error_logs.append(result_text)
            else:
                # 結果をCSVとしてパースしてリストに追加
                try:
                    # 文字列をDataFrameに変換（ヘッダーなし前提）
                    df_temp = pd.read_csv(io.StringIO(result_text), header=None, names=ICS_COLUMNS, dtype=str).fillna("")
                    new_data_rows.append(df_temp)
                except Exception as e:
                    error_logs.append(f"データ変換エラー: {target['name']} - {e}")

            progress_bar.progress((i + 1) / total_targets)
            time.sleep(5.0) # 待機

        # --- 3. データの結合と統合（名寄せ） ---
        if new_data_rows:
            # 今回のデータを結合
            current_batch_df = pd.concat(new_data_rows, ignore_index=True)
            
            # 過去のデータ(Session State)と結合
            st.session_state.master_dataframe = pd.concat([st.session_state.master_dataframe, current_batch_df], ignore_index=True)
            
            # ★データ統合処理（ここが重要）
            # 氏名（姓・名）と個人コードでグループ化し、空白を埋めるようにマージする
            # 'first' は最初の非欠損値を使うが、ここでは簡易的に文字列の最大値（空白より文字がある方）を採用
            # ※より厳密にするには個人コードをキーにするのがベスト
            
            # データフレーム内の全データを文字列型にしておく
            df_merged = st.session_state.master_dataframe.astype(str)
            
            # 空文字をNaNにして、bfill/ffillしやすくする手もあるが、
            # ここでは「個人コード」と「氏名」が同じなら、情報を合体させる
            # groupby().max() を使うと、空欄 "" よりも "10000" 等の値が優先される性質を利用
            
            if not df_merged.empty:
                # キーになる列（個人コードがある場合はそれを優先、なければ氏名）
                # ここでは簡易的に「氏名（姓）」「氏名（名）」で名寄せします
                df_merged = df_merged.groupby(['氏名（姓）', '氏名（名）'], as_index=False).max()
                
                # 並び順をICSの定義通りに戻す（groupbyで崩れることがあるため）
                df_merged = df_merged[ICS_COLUMNS]
                
                # 統合結果を保存
                st.session_state.master_dataframe = df_merged

            st.success(f"解析完了！ 現在、合計 {len(st.session_state.master_dataframe)} 名分のデータがあります。")
            
        else:
            if not error_logs:
                st.warning("有効なデータが見つかりませんでした。")

        # エラー表示
        if error_logs:
            st.error("以下のエラーが発生しました:")
            for err in error_logs:
                st.text(err)

    # --- 4. 結果のダウンロード（常に表示） ---
    if not st.session_state.master_dataframe.empty:
        st.write("---")
        st.subheader("📊 現在の統合データ")
        st.dataframe(st.session_state.master_dataframe)
        
        # CSV化
        final_csv_content = "給与,12,月分,FMT,1,DBVER\n"
        final_csv_content += "テンプレ,タイプ,SL=4\n"
        final_csv_content += "個人コード,区分コード,氏名（姓）,氏名（名）,氏名フリガナ（姓）,氏名フリガナ（名）,性別,入社年月日,退職年月日,郵便番号,住所1,住所2,1月支給,2月支給,3月支給,4月支給,5月支給,6月支給,7月支給,8月支給,9月支給,10月支給,11月支給,12月支給,1月社保,2月社保,3月社保,4月社保,5月社保,6月社保,7月社保,8月社保,9月社保,10月社保,11月社保,12月社保,1月税額,2月税額,3月税額,4月税額,5月税額,6月税額,7月税額,8月税額,9月税額,10月税額,11月税額,12月税額,1月賞与,2月賞与,3月賞与,4月賞与,5月賞与,6月賞与,7月賞与,8月賞与,9月賞与,10月賞与,11月賞与,12月賞与\n"
        
        # データ行を追加
        data_rows = st.session_state.master_dataframe.to_csv(index=False, header=False)
        final_csv_content += data_rows

        try:
            csv_bytes = final_csv_content.encode("cp932", errors="ignore")
            st.download_button(
                label="⬇️ 統合されたICS用CSVをダウンロード",
                data=csv_bytes,
                file_name="ics_import_master.csv",
                mime="text/csv"
            )
        except Exception as e:
            st.error(f"文字コード変換エラー: {e}")

elif not api_key:
    st.info("👈 左のサイドバーにGoogle APIキーを入力してください。")
