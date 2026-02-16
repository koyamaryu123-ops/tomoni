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
**Excelの全シート読み取り、データの自動統合（名寄せ）、追加読み込みに対応しています。**
""")

# ==========================================
# セッション状態の初期化
# ==========================================
if 'accumulated_rows' not in st.session_state:
    st.session_state.accumulated_rows = []
# 処理完了後のCSVデータを保持するための変数
if 'final_csv_data' not in st.session_state:
    st.session_state.final_csv_data = None

# サイドバーでAPIキー入力
api_key = st.sidebar.text_input("AIzaSyDCRsVPD7krj2iYrwrogh37RCsplx8S5lc", type="password")
if api_key:
    genai.configure(api_key=api_key)

# サイドバーにリセットボタン配置
if st.sidebar.button("データをリセット"):
    st.session_state.accumulated_rows = []
    st.session_state.final_csv_data = None
    st.rerun()

# ==========================================
# エリア1: 通常の給与データアップロード
# ==========================================
st.subheader("① 給与明細・台帳データのアップロード")
uploaded_files_main = st.file_uploader(
    "いつもの給与データはこちら", 
    type=["xlsx", "xls", "csv", "png", "jpg", "jpeg", "pdf"], 
    accept_multiple_files=True,
    key="uploader_main"
)

# ==========================================
# エリア2: 給与申告書（追加データ）アップロード
# ==========================================
st.subheader("② 給与申告書（追加修正用）のアップロード")
st.info("※ここでアップロードしたデータは、上のデータに追加・統合されます。")
uploaded_files_sub = st.file_uploader(
    "給与申告書など、追加したいファイルはこちら", 
    type=["xlsx", "xls", "csv", "png", "jpg", "jpeg", "pdf"], 
    accept_multiple_files=True,
    key="uploader_sub"
)

def process_single_file(content, filename, file_type="text"):
    """1つのファイルをAIに解析させ、データ行(CSV)だけを取り出す"""
    
    model = genai.GenerativeModel('gemini-2.5-flash') 
    
    # プロンプト（変更なし）
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
        if file_type == "pdf":
            pdf_data = {'mime_type': 'application/pdf', 'data': content}
            response = model.generate_content([prompt_text, pdf_data])
        elif file_type == "image":
            image = Image.open(io.BytesIO(content))
            response = model.generate_content([prompt_text, image])
        else:
            response = model.generate_content(prompt_text + f"\n\n【ファイル名: {filename} のデータ】\n{content}")

        raw_text = response.text.replace("```csv", "").replace("```", "").strip()
        return raw_text
        
    except Exception as e:
        return f"ERROR: {filename} の解析に失敗: {e}"

def merge_rows(all_rows):
    """収集したCSV行データをDataFrame化し、個人コードと氏名で名寄せ（統合）を行う関数"""
    if not all_rows:
        return []

    columns = [
        "個人コード","区分コード","氏名（姓）","氏名（名）","氏名フリガナ（姓）","氏名フリガナ（名）","性別","入社年月日","退職年月日","郵便番号","住所1","住所2",
        "1月支給","2月支給額","3月支給額","4月支給額","5月支給額","6月支給額","7月支給額","8月支給額","9月支給額","10月支給額","11月支給額","12月支給額",
        "1月社会保険料","2月社会保険料","3月社会保険料","4月社会保険料","5月社会保険料","6月社会保険料","7月社会保険料","8月社会保険料","9月社会保険料","10月社会保険料","11月社会保険料","12月社会保険料",
        "1月所得税","2月所得税","3月所得税","4月所得税","5月所得税","6月所得税","7月所得税","8月所得税","9月所得税","10月所得税","11月所得税","12月所得税",
        "1月賞与","2月賞与","3月賞与","4月賞与","5月賞与","6月賞与","7月賞与","8月賞与","9月賞与","10月賞与","11月賞与","12月賞与"
    ]

    data_list = []
    for row_str in all_rows:
        split_row = [x.strip().replace('"', '') for x in row_str.split(',')]
        if len(split_row) < len(columns):
            split_row += [""] * (len(columns) - len(split_row))
        else:
            split_row = split_row[:len(columns)]
        data_list.append(split_row)

    df = pd.DataFrame(data_list, columns=columns)
    df = df.replace(r'^\s*$', None, regex=True)

    group_keys = ['個人コード', '氏名（姓）', '氏名（名）']
    
    # 統合実行: 同じ人のデータがあれば、後から来たデータの空白でない部分で埋めるイメージ
    # groupby().first() は「最初の有効な値」を取るため、順序によっては上書きされない場合があるが、
    # 基本的には情報が集約される
    df_merged = df.groupby(group_keys, as_index=False).first()
    df_merged = df_merged.fillna("")

    merged_rows = []
    for _, row in df_merged.iterrows():
        row_str = ",".join([str(x) for x in row.values])
        merged_rows.append(row_str)
        
    return merged_rows

# ==========================================
# 実行ボタン
# ==========================================
# アップロードされた全ファイルをまとめる
all_uploaded_files = []
if uploaded_files_main:
    all_uploaded_files.extend(uploaded_files_main)
if uploaded_files_sub:
    all_uploaded_files.extend(uploaded_files_sub)

if all_uploaded_files and api_key:
    
    if st.button(f"全 {len(all_uploaded_files)} 件のファイルを一括解析・統合", type="primary"):
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        error_logs = []
        
        current_batch_rows = []
        total_files = len(all_uploaded_files)
        
        # --- ループ処理開始 ---
        for i, file in enumerate(all_uploaded_files):
            status_text.text(f"解析中 ({i+1}/{total_files}): {file.name} ...")
            
            files_to_process = []
            
            try:
                if file.name.endswith(('.xlsx', '.xls')):
                    excel_file = pd.ExcelFile(file)
                    for sheet_name in excel_file.sheet_names:
                        df = pd.read_excel(excel_file, sheet_name=sheet_name)
                        if not df.empty:
                            csv_text = df.to_csv(index=False)
                            files_to_process.append((csv_text, f"{file.name} [{sheet_name}]", "text"))
                            
                elif file.name.endswith('.csv'):
                    df = pd.read_csv(file)
                    content = df.to_csv(index=False)
                    files_to_process.append((content, file.name, "text"))
                    
                elif file.name.endswith('.pdf'):
                    content = file.read()
                    files_to_process.append((content, file.name, "pdf"))
                    
                elif file.name.endswith(('.png', '.jpg', '.jpeg')):
                    content = file.read()
                    files_to_process.append((content, file.name, "image"))

                # AI処理呼び出し
                for content, fname, ftype in files_to_process:
                    result_csv_text = process_single_file(content, fname, ftype)
                    
                    if result_csv_text.startswith("ERROR"):
                        error_logs.append(result_csv_text)
                    else:
                        rows = [r for r in result_csv_text.split('\n') if r.strip()]
                        current_batch_rows.extend(rows)

            except Exception as e:
                error_logs.append(f"ERROR: {file.name} の読み込み失敗: {e}")

            progress_bar.progress((i + 1) / total_files)

        # --- 結果をセッションに追記 ---
        if current_batch_rows:
            st.session_state.accumulated_rows.extend(current_batch_rows)
            st.success(f"今回 {len(current_batch_rows)} 行のデータを抽出しました。")
        
        status_text.text("データの統合処理（名寄せ）を行っています...")

        # --- 統合処理と結果保存 ---
        if st.session_state.accumulated_rows:
            merged_rows = merge_rows(st.session_state.accumulated_rows)

            # CSVデータ作成
            final_csv_content = "給与,12,月分,FMT,1,DBVER\n"
            final_csv_content += "テンプレ,タイプ,SL=4\n"
            ics_header_row = "個人コード,区分コード,氏名（姓）,氏名（名）,氏名フリガナ（姓）,氏名フリガナ（名）,性別,入社年月日,退職年月日,郵便番号,住所1,住所2,1月支給,2月支給,3月支給,4月支給,5月支給,6月支給,7月支給,8月支給,9月支給,10月支給,11月支給,12月支給,1月社会保険料,2月社会保険料,3月社会保険料,4月社会保険料,5月社会保険料,6月社会保険料,7月社会保険料,8月社会保険料,9月社会保険料,10月社会保険料,11月社会保険料,12月社会保険料,1月所得税,2月所得税,3月所得税,4月所得税,5月所得税,6月所得税,7月所得税,8月所得税,9月所得税,10月所得税,11月所得税,12月所得税,1月賞与,2月賞与,3月賞与,4月賞与,5月賞与,6月賞与,7月賞与,8月賞与,9月賞与,10月賞与,11月賞与,12月賞与"
            final_csv_content += ics_header_row + "\n"
            final_csv_content += "\n".join(merged_rows)

            # ★重要: セッションに保存して消えないようにする
            try:
                st.session_state.final_csv_data = final_csv_content.encode("cp932", errors="ignore")
            except:
                # CP932変換エラー時はUTF-8で逃げる（文字化けリスクあるがエラー落ち回避）
                st.session_state.final_csv_data = final_csv_content.encode("utf-8")

            # 完了通知
            st.toast("すべての解析が完了しました！", icon="🎉")
            st.balloons()
            
            # 音で通知（ブラウザの仕様によっては自動再生されない場合があります）
            # フリー素材の通知音URLなどを設定
            st.markdown("""
                <audio autoplay>
                <source src="https://assets.mixkit.co/active_storage/sfx/2869/2869-preview.mp3" type="audio/mpeg">
                </audio>
                """, unsafe_allow_html=True)
        else:
            st.warning("データがありません。")

        if error_logs:
            st.error("エラーが発生しました:")
            for err in error_logs:
                st.text(err)

# ==========================================
# ダウンロードエリア（処理後も常に表示）
# ==========================================
# セッションにデータがあれば、いつでもダウンロードボタンを表示する
if st.session_state.final_csv_data:
    st.divider()
    st.success("▼ データの準備ができています（時間が経っても消えません）")
    
    st.download_button(
        label="統合データをダウンロード (CSV)",
        data=st.session_state.final_csv_data,
        file_name="ics_import_merged.csv",
        mime="text/csv"
    )

elif not api_key:
    st.info("👈 左のサイドバーにGoogle APIキーを入力してください。")
