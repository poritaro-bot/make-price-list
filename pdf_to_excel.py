import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import unicodedata

def format_text(text):
    """英数字を半角に、半角カタカナを全角に変換する関数"""
    if pd.isna(text) or text is None:
        return ""
    # NFKC正規化：全角英数字→半角、半角カナ→全角カナ に自動変換してくれます
    return unicodedata.normalize('NFKC', str(text))

def clean_price(value):
    """金額から「¥」や「,（カンマ）」などの不要な文字を取り除いて数値化する関数"""
    if not value: return ""
    cleaned = re.sub(r'[^\d.]', '', str(value))
    if cleaned:
        return float(cleaned)
    return ""

def highlight_duplicates(df):
    """品番が重複している行全体を薄紫に色付けする関数"""
    # keep=False で、重複しているものは「すべて」Trueになります
    is_dupe = df.duplicated(subset=['品番'], keep=False)
    # 品番が空欄のものは色付け対象から外す
    is_not_empty = df['品番'].astype(str).str.strip() != ""
    mask = is_dupe & is_not_empty
    
    # 元のデータと同じ形のエクセルの「書式（スタイル）用シート」を作るイメージ
    style_df = pd.DataFrame('', index=df.index, columns=df.columns)
    
    # 重複している行に薄紫（#E6E6FA）を設定
    style_df[mask] = 'background-color: #E6E6FA'
    return style_df

# --- 画面の見た目を作る部分 ---
st.set_page_config(page_title="PDF→Excel変換アプリ", layout="centered")
st.title("📄 PDF価格表 → Excel変換ツール")
st.write("PDFをアップロードして、各列の見出しを指定するだけで指定フォーマットのExcelに変換します。")

# 1. ファイルアップロード
st.markdown("### 1. PDFファイルの選択")
uploaded_file = st.file_uploader("ここにPDFファイルをドラッグ＆ドロップ、または選択してください", type="pdf")

# 2. 列名マッピング
st.markdown("### 2. PDF内の列名（ヘッダー）を指定")
st.write("※空白にした項目は、エクセル上でも空白になります。")
col1, col2 = st.columns(2)
with col1:
    col_pn = st.text_input("A列 (品番) に入れるPDFの列名", value="商品コード")
    col_price = st.text_input("C列 (定価) に入れるPDFの列名", value="標準価格")
    col_name = st.text_input("E列 (商品名) に入れるPDFの列名", value="品名")
with col2:
    col_bm = st.text_input("B列 (BM) に入れるPDFの列名", value="BSC")
    col_cost = st.text_input("D列 (仕切り) に入れるPDFの列名", value="卸単価")

# 3. オプション設定
st.markdown("### 3. オプション")
apply_tax = st.checkbox("仕切り価格(D列)に 1.06 を掛けて計算する", value=True)

# 実行ボタン
st.markdown("---")
if uploaded_file is not None:
    if st.button("✨ Excelに変換する", type="primary", use_container_width=True):
        with st.spinner("PDFを解析中... 少々お待ちください。"):
            target_cols = {
                "品番": col_pn.strip(),
                "BM": col_bm.strip(),
                "定価": col_price.strip(),
                "仕切り": col_cost.strip(),
                "商品名": col_name.strip()
            }

            all_data = []
            try:
                # PDFの解析処理
                with pdfplumber.open(uploaded_file) as pdf:
                    for page in pdf.pages:
                        tables = page.extract_tables()
                        for table in tables:
                            if not table: continue
                            df_table = pd.DataFrame(table)
                            header_idx = -1
                            col_indices = { "品番": -1, "BM": -1, "定価": -1, "仕切り": -1, "商品名": -1 }

                            # 見出し探し
                            for i, row in df_table.iterrows():
                                row_list = [str(cell).replace('\n', '') if cell else "" for cell in row.tolist()]
                                match_count = 0
                                temp_indices = { "品番": -1, "BM": -1, "定価": -1, "仕切り": -1, "商品名": -1 }
                                for key, target_name in target_cols.items():
                                    if target_name:
                                        for col_i, cell_val in enumerate(row_list):
                                            if target_name in cell_val:
                                                temp_indices[key] = col_i
                                                match_count += 1
                                                break
                                if match_count > 0:
                                    header_idx = i
                                    col_indices = temp_indices
                                    break

                            if header_idx == -1: continue

                            # データ行の抽出
                            for i, row in df_table.iloc[header_idx + 1:].iterrows():
                                row_list = [str(cell).replace('\n', '') if cell else "" for cell in row.tolist()]
                                if all(cell.strip() == "" for cell in row_list): continue

                                # ここで format_text() を使って全角・半角の文字整形を行います
                                extracted = {
                                    "品番": format_text(row_list[col_indices["品番"]]) if col_indices["品番"] != -1 and col_indices["品番"] < len(row_list) else "",
                                    "BM": format_text(row_list[col_indices["BM"]]) if col_indices["BM"] != -1 and col_indices["BM"] < len(row_list) else "",
                                    "定価": format_text(row_list[col_indices["定価"]]) if col_indices["定価"] != -1 and col_indices["定価"] < len(row_list) else "",
                                    "仕切り": format_text(row_list[col_indices["仕切り"]]) if col_indices["仕切り"] != -1 and col_indices["仕切り"] < len(row_list) else "",
                                    "商品名": format_text(row_list[col_indices["商品名"]]) if col_indices["商品名"] != -1 and col_indices["商品名"] < len(row_list) else ""
                                }

                                if apply_tax and extracted["仕切り"]:
                                    val = clean_price(extracted["仕切り"])
                                    if val != "":
                                        extracted["仕切り"] = round(val * 1.06)

                                all_data.append(extracted)

                if not all_data:
                    st.error("抽出できるデータが見つかりませんでした。PDFの形式や入力した列名を確認してください。")
                else:
                    # Excelデータの作成
                    final_df = pd.DataFrame(all_data)
                    final_df = final_df[["品番", "BM", "定価", "仕切り", "商品名"]]

                    # 重複ハイライトのスタイルを適用
                    styled_df = final_df.style.apply(highlight_duplicates, axis=None)

                    # メモリ上でExcelデータを作成
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        styled_df.to_excel(writer, index=False, sheet_name='PriceList')
                    excel_data = output.getvalue()

                    st.success("🎉 変換が完了しました！")
                    
                    # ダウンロードボタンの表示
                    file_name = uploaded_file.name.replace('.pdf', '_変換結果.xlsx')
                    st.download_button(
                        label="📥 Excelファイルをダウンロード",
                        data=excel_data,
                        file_name=file_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )

            except Exception as e:
                st.error(f"エラーが発生しました: {e}")