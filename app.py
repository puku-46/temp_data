import io
import openpyxl
import pandas as pd
import streamlit as st

st.title("室温データ抽出・集計ツール")

# 1. ファイルのドラッグ＆ドロップ（アップロード）
uploaded_file = st.file_uploader(
    "ACC.xlsx または 室温データのCSVファイルをドラッグ＆ドロップしてください",
    type=["xlsx", "xls", "csv", "CSV"],
)

if uploaded_file is not None:
    try:
        # ファイル形式と文字コードの判定処理
        if uploaded_file.name.lower().endswith(".csv"):
            df_acc = None
            encodings = ["utf-8-sig", "cp932", "shift_jis"]

            # アップロードされたファイルのバイトデータを一時保存
            file_bytes = uploaded_file.getvalue()

            for enc in encodings:
                try:
                    df_acc = pd.read_csv(io.BytesIO(file_bytes), encoding=enc)
                    break
                except UnicodeDecodeError:
                    continue

            if df_acc is None:
                df_acc = pd.read_csv(
                    io.BytesIO(file_bytes), encoding="cp932", errors="replace"
                )
        else:
            df_acc = pd.read_excel(uploaded_file, header=0)

        # 2. 部屋・項目の選択（プルダウンメニュー）
        room_columns = list(df_acc.columns[2:])

        if not room_columns:
            st.error("データ列が見つかりませんでした。ファイル形式を確認してください。")
        else:
            target_col = st.selectbox("抽出したい部屋・項目を選択してください:", room_columns)

            if target_col:
                # 3. データの抽出（日付、時刻、指定された列）
                extracted_data = []
                for _, row in df_acc.iterrows():
                    file_name = str(row.iloc[0]).strip()
                    time_val = str(row.iloc[1]).strip()
                    val = row[target_col]

                    # 時刻の判定
                    if time_val in ["8時", "08時", "8:00", "08:00"]:
                        time_key = "8時"
                    elif time_val in ["17時", "17:00"]:
                        time_key = "17時"
                    else:
                        continue

                    extracted_data.append(
                        {"日付": file_name, "時刻": time_key, "温度": val}
                    )

                df_extracted = pd.DataFrame(extracted_data)

                # 4. ピボット展開とレイアウト組み立て
                if not df_extracted.empty:
                    df_pivot = df_extracted.pivot(
                        index="日付", columns="時刻", values="温度"
                    )

                    cols_order = [c for c in ["8時", "17時"] if c in df_pivot.columns]
                    df_pivot = df_pivot[cols_order]
                    df_pivot.reset_index(inplace=True)

                    output_rows = []
                    output_rows.append(["部屋", target_col, ""])
                    output_rows.append(["", "8時", "17時"])

                    for _, r in df_pivot.iterrows():
                        d = r["日付"]
                        v8 = r.get("8時", "")
                        v17 = r.get("17時", "")

                        try:
                            v8 = (
                                float(v8)
                                if pd.notna(v8) and str(v8).strip() != ""
                                else ""
                            )
                        except ValueError:
                            v8 = ""

                        try:
                            v17 = (
                                float(v17)
                                if pd.notna(v17) and str(v17).strip() != ""
                                else ""
                            )
                        except ValueError:
                            v17 = ""

                        output_rows.append([d, v8, v17])

                    df_out = pd.DataFrame(output_rows)

                    # 画面上にプレビュー表示
                    st.subheader("抽出データプレビュー")
                    st.dataframe(df_out)

                    # 5. Excelファイル（メモリ上）の生成とセル書式設定
                    output_buffer = io.BytesIO()
                    with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
                        df_out.to_excel(writer, index=False, header=False)

                        worksheet = writer.sheets["Sheet1"]
                        for row in range(3, len(output_rows) + 1):
                            for col in [2, 3]:
                                cell = worksheet.cell(row=row, column=col)
                                if isinstance(cell.value, (int, float)):
                                    cell.number_format = "0.0"

                    output_buffer.seek(0)

                    # 6. ダウンロードボタンの表示
                    st.download_button(
                        label="集計結果のExcelファイルをダウンロード",
                        data=output_buffer,
                        file_name=f"cal_output_{target_col}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                else:
                    st.warning("指定された条件（8時・17時）に合致するデータが見つかりませんでした。")

    except Exception as e:
        st.error(f"エラーが発生しました: {e}")
