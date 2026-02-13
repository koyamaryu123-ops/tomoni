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

st.title("📄 ICS年末調整データ作成 AIツール (PDF対応・一括処理版)")
st.markdown("""
複数の Excel、CSV、**給与明細の画像・PDF** をまとめてアップロードできます。
AIが全てのデータを読み取り、**1つのICS用CSVファイル** に統合して出力します。
""")

# サイドバーでAPIキー入力
api_key = st.sidebar.text_input("Google API Keyを入力", type="password")
if api_key:
    genai.configure(api_key=api_key)

# ==========================================
# ファイルアップロード (複数対応)
# ==========================================
uploaded_files = st.file_uploader(
    "ファイルをまとめてドラッグ＆ドロップしてください", 
    type=["xlsx", "xls", "csv", "png", "jpg", "jpeg", "pdf"], 
    accept_multiple_files=True
)

def process_single_file(content, filename, file_type="text"):
    """1つのファイルをAIに解析させ、データ行(CSV)だけを取り出す（リトライ機能付き）"""
    
    model = genai.GenerativeModel('gemini-2.5-flash') 
    
    # プロンプト（元の内容を維持）
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

    # ★変更点：最大3回までリトライする処理を追加
    max_retries = 3
    
    for attempt in range(max_retries):
        try:
            # ファイルタイプに応じた処理
            if file_type == "pdf":
                pdf_data = {'mime_type': 'application/pdf', 'data': content}
                response = model.generate_content([prompt_text, pdf_data])
            elif file_type == "image":
                image = Image.open(io.BytesIO(content))
                response = model.generate_content([prompt_text, image])
            else:
                response = model.generate_content(prompt_text + f"\n\n【ファイル名: {filename} のデータ】\n{content}")

            # 成功したらループを抜けて結果を返す
            raw_text = response.text.replace("```csv", "").replace("```", "").strip()
            return raw_text
        
        except Exception as e:
            error_msg = str(e)
            # クォータ制限のエラーが出たら待機してリトライ
            if "429" in error_msg or "quota" in error_msg.lower() or "resource" in error_msg.lower():
                if attempt < max_retries - 1:
                    wait_time = 20 # 20秒待機
                    st.toast(f"⚠️ API制限のため待機中... ({filename}: {attempt+1}/{max_retries})")
                    time.sleep(wait_time)
                    continue
                else:
                    return f"ERROR: API制限により解析失敗 ({filename})"
            else:
                # その他のエラーは即終了
                return f"ERROR: {filename} の解析に失敗: {e}"

# ==========================================
# 実行ボタン
# ==========================================
if uploaded_files and api_key:
    if st.button(f"選択した {len(uploaded_files)} 件のファイルを一括解析", type="primary"):
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        all_csv_rows = [] 
        error_logs = []
        total_files = len(uploaded_files)
        
        # --- ループ処理開始 ---
        for i, file in enumerate(uploaded_files):
            status_text.text(f"解析中 ({i+1}/{total_files}): {file.name} ...")
            
            # 前処理
            input_content = None
            file_type = "text"
            
            try:
                if file.name.endswith('.pdf'):
                    input_content = file.read()
                    file_type = "pdf"
                elif file.name.endswith(('.png', '.jpg', '.jpeg')):
                    input_content = file.read()
                    file_type = "image"
                elif file.name.endswith('.csv'):
                    df = pd.read_csv(file)
                    input_content = df.to_csv(index=False)
                    file_type = "text"
                elif file.name.endswith(('.xlsx', '.xls')):
                    df = pd.read_excel(file)
                    input_content = df.to_csv(index=False)
                    file_type = "text"
                
                # AI処理呼び出し
                result_csv_text = process_single_file(input_content, file.name, file_type)
                
                if result_csv_text.startswith("ERROR"):
                    error_logs.append(result_csv_text)
                else:
                    rows = [r for r in result_csv_text.split('\n') if r.strip()]
                    all_csv_rows.extend(rows)
                    
            except Exception as e:
                error_logs.append(f"ERROR: {file.name} の読み込み失敗: {e}")

            # 進捗バー更新
            progress_bar.progress((i + 1) / total_files)
            
            # ★変更点：制限回避のため待機時間を10秒に延長（安全策）
            time.sleep(10.0) 

        status_text.text("全ファイルの解析が完了しました！ CSVを作成しています...")

        # --- 結果の結合と出力 ---
        if all_csv_rows:
            final_csv_content = "給与,12,月分,FMT,1,DBVER\n"
            final_csv_content += "テンプレ,タイプ,SL=4\n"
            
            ics_header_row = "個人コード,区分コード,氏名（姓）,氏名（名）,氏名フリガナ（姓）,氏名フリガナ（名）,性別,入社年月日,退職年月日,郵便番号,住所1,住所2,1月支給,2月支給,3月支給,4月支給,5月支給,6月支給,7月支給,8月支給,9月支給,10月支給,11月支給,12月支給,1月社会保険料,2月社会保険料,3月社会保険料,4月社会保険料,5月社会保険料,6月社会保険料,7月社会保険料,8月社会保険料,9月社会保険料,10月社会保険料,11月社会保険料,12月社会保険料,1月所得税,2月所得税,3月所得税,4月所得税,5月所得税,6月所得税,7月所得税,8月所得税,9月所得税,10月所得税,11月所得税,12月所得税,1月賞与,2月賞与,3月賞与,4月賞与,5月賞与,6月賞与,7月賞与,8月賞与,9月賞与,10月賞与,11月賞与,12月賞与"
            final_csv_content += ics_header_row + "\n"
            final_csv_content += "\n".join(all_csv_rows)

            st.success(f"合計 {len(all_csv_rows)} 人分のデータを抽出・結合しました！")
            st.text_area("CSVプレビュー (結合結果)", final_csv_content, height=200)

            try:
                csv_bytes = final_csv_content.encode("cp932", errors="ignore")
                st.download_button(
                    label="統合されたICS用CSVをダウンロード",
                    data=csv_bytes,
                    file_name="ics_import_merged.csv",
                    mime="text/csv"
                )
            except Exception as e:
                st.error(f"文字コード変換エラー: {e}")
        else:
            st.warning("有効なデータが1件も抽出できませんでした。")

        if error_logs:
            st.error("一部のファイルでエラーが発生しました:")
            for err in error_logs:
                st.text(err)

elif not api_key:
    st.info("👈 左のサイドバーにGoogle APIキーを入力してください。")
