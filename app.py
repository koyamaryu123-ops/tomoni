import streamlit as st
import pandas as pd
import google.generativeai as genai
import json
import io
import csv
from PIL import Image

# ==========================================
# 設定: ICSヘッダー定義
# ==========================================
ICS_HEADERS = [
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

# ==========================================
# ページ設定
# ==========================================
st.set_page_config(page_title="ICS年末調整データ作成ツール", layout="wide")

st.title("📄 ICS年末調整データ作成 AIツール")
st.markdown("""
Excel、CSV、または **給与明細の画像** をアップロードすると、AIが自動で読み取り、
ICSシステムに取り込める形式（CSV）に変換します。
""")

# サイドバーでAPIキー入力
api_key = st.sidebar.text_input("Google API Keyを入力", type="password")
if api_key:
    genai.configure(api_key=api_key)

# ==========================================
# ファイルアップロード
# ==========================================
uploaded_file = st.file_uploader("ファイルをドラッグ＆ドロップしてください", type=["xlsx", "xls", "csv", "png", "jpg", "jpeg"])

def process_with_ai(content, mime_type, is_image=False):
    """AIにデータを投げてJSON化する"""
    
    # モデル設定（エラーが出る場合は gemini-1.5-flash 等に変更してください）
    model = genai.GenerativeModel('gemini-2.5-flash') 
    
    # ★修正ポイント：プロンプトから「CSV出力ルール」を削除し、JSON出力に特化させました
    prompt_text = """
    あなたは給与計算のプロフェッショナルです。
    提供されたデータ（テキストまたは画像）から、従業員の給与情報を読み取り、
    **必ず以下のJSON形式のみ** で出力してください。

    【重要ルール】
    1. 金額の「円」やカンマ「,」はすべて削除し、半角数字のみにする。
    2. データが存在しない月や項目は、空文字 "" (引用符の中身なし) にする。
    3. 氏名は「姓」と「名」に分割する。
    4. 日付は YYYY/MM/DD 形式に統一する。
    5. **出力は純粋なJSONテキストのみ返すこと。** Markdownタグ（```json 等）や挨拶文は一切不要です。

    【抽出項目キー（JSONのキー名）】
    personal_code, section_code, last_name, first_name, last_name_kana, first_name_kana,
    gender (1=男, 2=女), hire_date, retire_date, zip_code, address1, address2,
    salary_1~12 (各月の支給), social_insurance_1~12 (各月の社保), tax_1~12 (各月の税), bonus_1~12 (各月の賞与)
    添付されたファイルから給与データを抽出し、ICS年末調整システムの入力形式に合わせて出力してください。一個入力がずれると全てずれてしまうことを加味し、正確に情報を読み取り、正確にファイルを作成しなさい。







    【出力ルール】

    1. 形式はcsvとし、ヘッダー(項目名)を必ず含めること。

    2. 列の並び：一行目に給与,12,月分,FMT,1,DBVER,二行目にテンプレ,タイプ,SL=4、3行目に個人コード,区分コード,氏名（姓）,氏名（名）,氏名フリガナ（姓）,氏名フリガナ（名）,性別,入社年月日,退職年月日,郵便番号,住所1,住所2,各月の支給額(N月分支給額).各月の社会保険料(N月分社会保険料),各月の所得税(N月分所得税),各月の賞与(N月分の賞与)

    3. 行の構成：それぞれ個人ごとに行を改行する。

    4. データがない項目は、項目名だけ書き、数字の場所は空欄(カンマのみ）にすること。

    5. 数字にカンマや「円」を含めないこと。

    6. 出力はcsvテキストのみ。余計な説明は一切不要。
    """

    try:
        if is_image:
            image = Image.open(content)
            response = model.generate_content([prompt_text, image])
        else:
            response = model.generate_content(prompt_text + f"\n\n【データ】\n{content}")

        # デバッグ用：AIの生の返答をログに出す（エラー時の確認用）
        print(f"DEBUG: AI Response:\n{response.text}")

        cleaned_json = response.text.replace("```json", "").replace("```", "").strip()
        return json.loads(cleaned_json)
        
    except Exception as e:
        # 画面にもエラー詳細を表示
        st.error(f"AI解析エラー: {e}")
        st.text("AIからの返答内容:")
        try:
            st.text(response.text) # AIが何を返したか確認できるようにする
        except:
            pass
        return []

# ==========================================
# 実行ボタン
# ==========================================
if uploaded_file and api_key:
    if st.button("AI解析スタート", type="primary"):
        with st.spinner("AIがデータを解析中..."):
            
            # ファイル形式ごとの前処理
            input_content = None
            is_image = False
            
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
                input_content = df.to_csv(index=False)
            elif uploaded_file.name.endswith(('.xlsx', '.xls')):
                df = pd.read_excel(uploaded_file)
                input_content = df.to_csv(index=False)
            else:
                input_content = uploaded_file
                is_image = True

            # AI解析実行
            data_list = process_with_ai(input_content, uploaded_file.type, is_image=is_image)
            
            if isinstance(data_list, dict): 
                data_list = [data_list]

            if data_list:
                st.success(f"{len(data_list)}件のデータを抽出しました！")
                
                # プレビュー表示
                preview_df = pd.DataFrame(data_list)
                st.dataframe(preview_df)

                # ICS形式 CSV生成処理
                output = io.StringIO()
                writer = csv.writer(output)
                
                # ICSヘッダー
                writer.writerow(["給与", "12", "月分", "FMT", "1", "DBVER"])
                writer.writerow(["テンプレ", "タイプ", "SL=4"])
                writer.writerow(ICS_HEADERS)

                for entry in data_list:
                    row = []
                    # 基本情報
                    row.append(entry.get("personal_code", ""))
                    row.append(entry.get("section_code", ""))
                    row.append(entry.get("last_name", ""))
                    row.append(entry.get("first_name", ""))
                    row.append(entry.get("last_name_kana", ""))
                    row.append(entry.get("first_name_kana", ""))
                    row.append(entry.get("gender", ""))
                    row.append(entry.get("hire_date", ""))
                    row.append(entry.get("retire_date", ""))
                    row.append(entry.get("zip_code", ""))
                    row.append(entry.get("address1", ""))
                    row.append(entry.get("address2", ""))
                    
                    def fmt(val): return str(val) if val and val != 0 else ""
                    
                    for key in ["salary", "social_insurance", "tax", "bonus"]:
                        for i in range(1, 13):
                            # JSONキーは salary_1, salary_2... となっている想定
                            row.append(fmt(entry.get(f"{key}_{i}")))
                            
                    writer.writerow(row)

                csv_data = output.getvalue().encode("cp932", errors="ignore")

                st.download_button(
                    label="ICS取込用CSVをダウンロード",
                    data=csv_data,
                    file_name="ics_import_data.csv",
                    mime="text/csv"
                )
            else:
                st.warning("データを抽出できませんでした。AIの応答を確認してください。")

elif not api_key:
    st.info("👈 左のサイドバーにGoogle APIキーを入力してください。")
