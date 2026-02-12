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

st.title("📄 ICS年末調整データ作成 AIツール (一括処理版)")
st.markdown("""
複数の Excel、CSV、**給与明細の画像** をまとめてアップロードできます。
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
    type=["xlsx", "xls", "csv", "png", "jpg", "jpeg"],
    accept_multiple_files=True  # ★ここが重要：複数選択OK
)

def process_single_file(content, filename, is_image=False):
    """1つのファイルをAIに解析させ、データ行(CSV)だけを取り出す"""
    
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

    try:
        if is_image:
            image = Image.open(content)
            response = model.generate_content([prompt_text, image])
        else:
            response = model.generate_content(prompt_text + f"\n\n【ファイル名: {filename} のデータ】\n{content}")

        # 結果のクリーニング
        raw_text = response.text.replace("```csv", "").replace("```", "").strip()
        return raw_text
        
    except Exception as e:
        return f"ERROR: {filename} の解析に失敗: {e}"

# ==========================================
# 実行ボタン
# ==========================================
if uploaded_files and api_key:
    if st.button(f"選択した {len(uploaded_files)} 件のファイルを一括解析", type="primary"):
        
        # 進捗バーの表示
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        all_csv_rows = [] # ここに全ファイルの解析結果(行)を貯める
        error_logs = []

        total_files = len(uploaded_files)
        
        # --- ループ処理開始 ---
        for i, file in enumerate(uploaded_files):
            status_text.text(f"解析中 ({i+1}/{total_files}): {file.name} ...")
            
            # 前処理
            input_content = None
            is_image = False
            
            try:
                if file.name.endswith('.csv'):
                    df = pd.read_csv(file)
                    input_content = df.to_csv(index=False)
                elif file.name.endswith(('.xlsx', '.xls')):
                    df = pd.read_excel(file)
                    input_content = df.to_csv(index=False)
                else:
                    input_content = file
                    is_image = True
                
                # AI処理呼び出し
                result_csv_text = process_single_file(input_content, file.name, is_image)
                
                if result_csv_text.startswith("ERROR"):
                    error_logs.append(result_csv_text)
                else:
                    # 空行を除去してリストに追加
                    rows = [r for r in result_csv_text.split('\n') if r.strip()]
                    all_csv_rows.extend(rows)
                    
            except Exception as e:
                error_logs.append(f"ERROR: {file.name} の読み込み失敗: {e}")

            # 進捗バー更新
            progress_bar.progress((i + 1) / total_files)
            time.sleep(0.5) # API制限回避のための少しの休憩

        status_text.text("全ファイルの解析が完了しました！ CSVを作成しています...")

        # --- 結果の結合と出力 ---
        if all_csv_rows:
            # 1. ICS固定ヘッダー
            final_csv_content = "給与,12,月分,FMT,1,DBVER\n"
            final_csv_content += "テンプレ,タイプ,SL=4\n"
            
            # 2. 項目名ヘッダー
            ics_header_row = "個人コード,区分コード,氏名（姓）,氏名（名）,氏名フリガナ（姓）,氏名フリガナ（名）,性別,入社年月日,退職年月日,郵便番号,住所1,住所2,1月支給,2月支給,3月支給,4月支給,5月支給,6月支給,7月支給,8月支給,9月支給,10月支給,11月支給,12月支給,1月社会保険料,2月社会保険料,3月社会保険料,4月社会保険料,5月社会保険料,6月社会保険料,7月社会保険料,8月社会保険料,9月社会保険料,10月社会保険料,11月社会保険料,12月社会保険料,1月所得税,2月所得税,3月所得税,4月所得税,5月所得税,6月所得税,7月所得税,8月所得税,9月所得税,10月所得税,11月所得税,12月所得税,1月賞与,2月賞与,3月賞与,4月賞与,5月賞与,6月賞与,7月賞与,8月賞与,9月賞与,10月賞与,11月賞与,12月賞与"
            final_csv_content += ics_header_row + "\n"
            
            # 3. AIが抽出したデータ行を全部結合
            final_csv_content += "\n".join(all_csv_rows)

            st.success(f"合計 {len(all_csv_rows)} 人分のデータを抽出・結合しました！")
            
            # プレビュー
            st.text_area("CSVプレビュー (結合結果)", final_csv_content, height=200)

            # ダウンロード
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

        # エラーがあったファイルの表示
        if error_logs:
            st.error("一部のファイルでエラーが発生しました:")
            for err in error_logs:
                st.text(err)

elif not api_key:
    st.info("👈 左のサイドバーにGoogle APIキーを入力してください。")
