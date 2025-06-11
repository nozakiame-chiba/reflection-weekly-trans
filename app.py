import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="リフレクション変換ツール", layout="centered")
st.title("週リフレクション 変換ツール")

uploaded_file = st.file_uploader("Excelファイル（元データ）をアップロードしてください", type="xlsx")

if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name=0)
    weeks = sorted(df['何週目の回答ですか。'].unique())

    st.success("✅ ファイル読み込み成功！")

    if st.button("変換を実行"):
        # ---- 「2_主要症候」 ----
        symptom_data = df.iloc[:, 4:41]
        symptom_names = symptom_data.columns.str.strip().str.replace(r'[\[\]]', '', regex=True)
        symptom_df = pd.DataFrame({'主要徴候': symptom_names})

        for week in weeks:
            week_rows = df['何週目の回答ですか。'] == week
            week_sum = df.loc[week_rows].iloc[:, 4:41].sum().values
            symptom_df[f'経験数【第{week}】'] = week_sum

        symptom_df['診療に参加した数'] = symptom_df[
            [col for col in symptom_df.columns if col.startswith('経験数【第')]
        ].sum(axis=1).astype(int)

        cols = ['主要徴候', '診療に参加した数'] + [c for c in symptom_df.columns if c.startswith('経験数【第')]
        symptom_df = symptom_df[cols]

        # ---- 「3_経験症例」 ----
        disease_df = df.iloc[:, 41:]
        disease_columns = disease_df.columns

        disease_info = []
        for col in disease_columns:
            match = re.findall(r'([^\[\]]+)\[([^\[\]]+)\]', col.strip())
            if match:
                category, disease = match[0]
                disease_info.append({'column': col, '分類': category.strip(), '疾患名': disease.strip()})

        records = []
        for item in disease_info:
            row = {
                '分類': item['分類'],
                '経験症例': item['疾患名']
            }
            total = 0
            for week in weeks:
                week_rows = df['何週目の回答ですか。'] == week
                count = df.loc[week_rows, item['column']].sum()
                row[f'経験数【第{week}】'] = count
                total += count
            row['診療に参加した数'] = total
            records.append(row)

        experience_df = pd.DataFrame(records)
        exp_cols = ['分類', '経験症例', '診療に参加した数'] + [c for c in experience_df.columns if c.startswith('経験数【第')]
        experience_df = experience_df[exp_cols]

        # Excel出力用バッファ
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            symptom_df.to_excel(writer, sheet_name='2_主要症候', index=False)
            experience_df.to_excel(writer, sheet_name='3_経験症例', index=False)

        st.success("✅ 変換完了！")
        st.download_button(
            label="📥 変換後ファイルをダウンロード",
            data=output.getvalue(),
            file_name="変換後_リフレクションA形式.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
