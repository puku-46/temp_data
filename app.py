import csv
import io
import re
import openpyxl
import pandas as pd
import streamlit as st

st.title("室温データ抽出・複数ファイル集計ツール")

# 1. 複数ファイルのドラッグ＆ドロップ（アップロード）
uploaded_files = st.file_uploader(
    "ACC.xlsx または 複数の室温データCSVファイルをまとめてドラッグ＆ドロップしてください",
    type=["xlsx", "xls", "csv", "CSV"],
    accept_multiple_files=True,
)

if uploaded_files:
  all_extracted_rows = []
  all_room_columns = set()
  file_dfs = []

  # アップロードされた各ファイルを処理
  for uploaded_file in uploaded_files:
    try:
      # ファイル名から不要な文字を除去して日付（例: 260801）を抽出
      raw_name = uploaded_file.name
      file_name_clean = (
          raw_name.replace(".CSV", "").replace(".csv", "").strip()
      )
      file_name_clean = re.sub(r"^室温_?", "", file_name_clean)
      file_name_clean = re.sub(r"_?室温$", "", file_name_clean)

      if raw_name.lower().endswith(".csv"):
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
          except Exception:
            continue

        if raw_rows is None:
          text_data = file_bytes.decode("cp932", errors="replace")
          f_io = io.StringIO(text_data)
          reader = csv.reader(f_io)
          raw_rows = list(reader)

        if len(raw_rows) < 7:
          continue

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
        df_file = pd.DataFrame(data_rows)
        if len(col_names) < len(df_file.columns):
          col_names += [
              f"Col_{i}" for i in range(len(col_names), len(df_file.columns))
          ]
        df_file.columns = col_names[: len(df_file.columns)]
        df_file.insert(0, "日付", file_name_clean)

        file_dfs.append(df_file)

        # 部屋・項目のリストを登録
        if len(df_file.columns) > 2:
          all_room_columns.update(df_file.columns[2:])

      else:
        # Excelファイル単体の場合
        df_file = pd.read_excel(uploaded_file, header=0)
        file_dfs.append(df_file)
        if len(df_file.columns) > 2:
          all_room_columns.update(df_file.columns[2:])

    except Exception as e:
      st.error(f"ファイル {uploaded_file.name} の処理中にエラー: {e}")

  if file_dfs and all_room_columns:
    # 部屋・項目一覧のソート
    sorted_rooms = sorted(list(all_room_columns))

    # 2. 抽出対象の部屋・項目を選択
    target_col = st.selectbox(
        "抽出したい部屋・項目を選択してください:", sorted_rooms
    )

    if target_col:
      extracted_data = []

      for df in file_dfs:
        if target_col not in df.columns:
          continue

        for _, row in df.iterrows():
          # 時刻セル（インデックス2または行全体から探索）
          time_val = str(row.iloc[2]).strip() if len(row) > 2 else ""

          time_key = None
          if time_val in ["08時", "8時", "08:00", "8:00"]:
            time_key = "8時"
          elif time_val in ["17時", "17:00"]:
            time_key = "17時"
          else:
            continue

          date_val = str(row.iloc[0]).strip()
          val = row[target_col]

          extracted_data.append(
              {"日付": date_val, "時刻": time_key, "温度": val}
          )

      df_extracted = pd.DataFrame(extracted_data)

      if not df_extracted.empty:
        # 重複削除＆日付順ソート
        df_extracted = df_extracted.drop_duplicates(subset=["日付", "時刻"])

        # ピボット展開（日付をインデックス、時刻を列へ）
        df_pivot = df_extracted.pivot(
            index="日付", columns="時刻", values="温度"
        )
        cols_order = [c for c in ["8時", "17時"] if c in df_pivot.columns]
        df_pivot = df_pivot[cols_order]
        df_pivot.reset_index(inplace=True)

        # 日付順でソート
        df_pivot.sort_values(by="日付", inplace=True)

        # 3. 集計レイアウトの組み立て
        output_rows = []
        output_rows.append(["部屋", target_col, ""])
        output_rows.append(["", "8時", "17時"])

        for _, r in df_pivot.iterrows():
          d = r["日付"]
          v8 = r.get("8時", "")
          v17 = r.get("17時", "")

          try:
            v8 = float(v8) if pd.notna(v8) and str(v8).strip() != "" else ""
          except ValueError:
            v8 = ""

          try:
            v17 = float(v17) if pd.notna(v17) and str(v17).strip() != "" else ""
          except ValueError:
            v17 = ""

          output_rows.append([d, v8, v17])

        df_out = pd.DataFrame(output_rows)

        # プレビュー表示
        st.subheader(f"【{target_col}】全ファイル集計結果プレビュー")
        st.dataframe(df_out)

        # 4. Excel生成＆数値フォーマット設定
        output_buffer = io.BytesIO()
        with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
          df_out.to_excel(writer, index=False, header=False)

          worksheet = writer.sheets["Sheet1"]
          for row_idx in range(3, len(output_rows) + 1):
            for col_idx in [2, 3]:
              cell = worksheet.cell(row=row_idx, column=col_idx)
              if isinstance(cell.value, (int, float)):
                cell.number_format = "0.0"

        output_buffer.seek(0)

        # 5. ダウンロードボタン
        st.download_button(
            label=f"【{target_col}】全日分まとめExcelをダウンロード",
            data=output_buffer,
            file_name=f"cal_summary_{target_col}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
      else:
        st.warning("指定された条件（8時・17時）に合致するデータが見つかりませんでした。")
