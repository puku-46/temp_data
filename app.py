import csv
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
      raw_rows = None
      encodings = ["cp932", "utf-8-sig", "shift_jis"]
      file_bytes = uploaded_file.getvalue()

      for enc in encodings:
        try:
          text_data = file_bytes.decode(enc)
          f_io = io.StringIO(text_data)
          reader = csv.reader(f_io)
          raw_rows = list(reader)
          break
        except UnicodeDecodeError:
          continue
        except Exception:
          continue

      if raw_rows is None:
        text_data = file_bytes.decode("cp932", errors="replace")
        f_io = io.StringIO(text_data)
        reader = csv.reader(f_io)
        raw_rows = list(reader)

      # ヘッダー情報の構築
      row2 = raw_rows[2] if len(raw_rows) > 2 else []
      row3 = raw_rows[3] if len(raw_rows) > 3 else []
      row4 = raw_rows[4] if len(raw_rows) > 4 else []

      curr_floor = ""
      curr_loc = ""
      col_names = []

      max_len = max(len(row2), len(row3), len(row4))
      for idx in range(max_len):
        f_val = (
            str(row2[idx]).strip()
            if idx < len(row2) and str(row2[idx]).strip()
            else ""
        )
        if f_val:
          curr_floor = f_val
        l_val = (
            str(row3[idx]).strip()
            if idx < len(row3) and str(row3[idx]).strip()
            else ""
        )
        if l_val:
          curr_loc = l_val
        r_val = (
            str(row4[idx]).strip()
            if idx < len(row4) and str(row4[idx]).strip()
            else ""
        )

        parts = [p for p in [curr_floor, curr_loc, r_val] if p]
        col_names.append("_".join(parts) if parts else f"Col_{idx}")

      data_rows = raw_rows[6:]
      df_acc = pd.DataFrame(data_rows)
      if len(col_names) < len(df_acc.columns):
        col_names += [
            f"Col_{i}" for i in range(len(col_names), len(df_acc.columns))
        ]
      df_acc.columns = col_names[: len(df_acc.columns)]

      file_name_clean = (
          uploaded_file.name.replace(".CSV", "").replace(".csv", "").strip()
      )
      df_acc.insert(0, "ファイル名", file_name_clean)
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
          time_val = str(row.iloc[1]).strip() if len(row) > 1 else ""

          # 時刻の判定（「8時」「08時」「8:00」などを柔軟にキャッチ）
          time_key = None
          if any(
              t in time_val for t in ["8時", "08時", "8:00", "08:00", "8時00"]
          ):
            time_key = "8時"
          elif any(t in time_val for t in ["17時", "17:00", "17時00"]):
            time_key = "17時"
          else:
            continue

          file_val = str(row.iloc[0]).strip()
          val = row[target_col]

          extracted_data.append({"日付": file_val, "時刻": time_key, "温度": val})

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
