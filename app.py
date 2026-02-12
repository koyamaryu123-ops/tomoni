import streamlit as st
import pandas as pd
import google.generativeai as genai
import io
from PIL import Image

# ==========================================
# ページ設定
# ==========================================
st.set_page_config(page_title="ICS年末調整データ作成ツール", layout="wide")

st.title("📄 ICS年末調整データ作成 AIツール")
st.markdown("""
Excel、CSV、または **給与明細の画像** をアップロードすると、
AIが **ICSシステム用のCSVデータ** を作成します。
""")

# サイドバーでAPIキー入力
api_key = st.sidebar.text_input("Google API Keyを入力", type="password")
if api_key:
    genai.configure(api_key=api_key)

# ==========================================
# ファイルアップロード
# ==========================================
uploaded_file = st.file_uploader("ファイルをドラッグ＆ドロップしてください", type=["xlsx", "xls", "csv", "png", "jpg", "jpeg"])

def process_with_ai(content, is_image=False):
    """AIにデータを投げてCSVテキストをもらう"""
    
    # モデル設定
    model = genai.GenerativeModel('gemini-2.5-flash') 
    
    # ★修正ポイント：AIに「CSVそのものを作れ」と指示します
    prompt_text = """
    あなたは給与計算のプロフェッショナルです。
    提供されたデータ（テキストまたは画像）から給与データを抽出し、
    **ICS年末調整システム用のCSV形式** で出力してください。

    【最重要：出力フォーマット】
    以下のCSV形式（カンマ区切り）のみを出力してください。Markdownタグや説明は一切不要です。

    行1: 給与,12,月分,FMT,1,DBVER
    行2: テンプレ,タイプ,SL=4
    行3: 個人コード,区分コード,氏名（姓）,氏名（名）,氏名フリガナ（姓）,氏名フリガナ（名）,性別,入社年月日,退職年月日,郵便番号,住所1,住所2,1月支給,2月支給,3月支給,4月支給,5月支給,6月支給,7月支給,8月支給,9月支給,10月支給,11月支給,12月支給,1月社保,2月社保,3月社保,4月社保,5月社保,6月社保,7月社保,8月社保,9月社保,10月社保,11月社保,12月社保,1月税額,2月税額,3月税額,4月税額,5月税額,6月税額,7月税額,8月税額,9月税額,10月税額,11月税額,12月税額,1月賞与,2月賞与,3月賞与,4月賞与,5月賞与,6月賞与,7月賞与,8月賞与,9月賞与,10月賞与,11月賞与,12月賞与
    行4以降: (抽出したデータ)

    【データ作成ルール】
    1. 金額の「円」やカンマは削除し、半角数字のみにする。
    2. データがない月や項目は、空欄のままカンマだけ打つこと（例: ,,,）。
       ※列ズレは絶対に許されません。カンマの数を正確に合わせてください。
    3. 氏名は姓と名に分割。
    4. 日付は YYYY/MM/DD 形式。
    5. 性別は 1=男, 2=女。
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

        # AIの返答（CSVテキスト）をそのまま返す
        return response.text
        
    except Exception as e:
        st.error(f"AI解析エラー: {e}")
        return ""

# ==========================================
# 実行ボタン
# ==========================================
if uploaded_file and api_key:
    if st.button("AI解析スタート", type="primary"):
        with st.spinner("AIがCSVを作成中..."):
            
            # 前処理
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
            raw_csv_text = process_with_ai(input_content, is_image=is_image)
            
            if raw_csv_text:
                # Markdownタグ（```csv ... ```）が付いていたら除去する処理
                cleaned_csv = raw_csv_text.replace("```csv", "").replace("```", "").strip()
                
                # 画面でプレビューを表示
                st.success("CSV作成完了！")
                st.text("▼ 作成されたCSVの中身（プレビュー）")
                st.text(cleaned_csv) # ここで中身を確認できます

                # ------------------------------------------
                # ダウンロード処理
                # ------------------------------------------
                # ICS用に Shift_JIS (cp932) に変換してダウンロードさせる
                try:
                    csv_data = cleaned_csv.encode("cp932", errors="ignore")
                    
                    st.download_button(
                        label="ICS取込用CSVをダウンロード",
                        data=csv_data,
                        file_name="ics_import_data.csv",
                        mime="text/csv"
                    )
                except Exception as e:
                    st.error(f"文字コード変換エラー: {e}")
            else:
                st.warning("データを生成できませんでした。")

elif not api_key:
    st.info("👈 左のサイドバーにGoogle APIキーを入力してください。")
