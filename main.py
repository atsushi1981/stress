import os
import pandas as pd

def extract_ab_columns(input_dir='./input', output_dir='./out'):
    # 出力フォルダの作成
    os.makedirs(output_dir, exist_ok=True)

    # 対象拡張子
    excel_extensions = ('.xlsx', '.xls')
    target_labels = {'A', 'B'}

    log = []

    # 入力ディレクトリ内のファイルを走査
    for filename in os.listdir(input_dir):
        if filename.endswith(excel_extensions):
            filepath = os.path.join(input_dir, filename)
            try:
                # Excelファイルを読み込み（全シート）
                xl = pd.ExcelFile(filepath)
                for sheet_name in xl.sheet_names:
                    try:
                        df = xl.parse(sheet_name, header=None)
                        if df.shape[0] < 2:
                            continue  # 2行目がなければスキップ
                        header_row = df.iloc[0]  # 2行目（インデックス1）

                        for i in range(len(header_row) - 1):
                            col_labels = {str(header_row[i]).strip(), str(header_row[i+1]).strip()}
                            if target_labels.issubset(col_labels):
                                # A, B列が隣接して見つかった場合
                                extracted_df = df.iloc[2:, [i, i+1]]  # 3行目以降（インデックス2以降）
                                extracted_df.columns = ['A', 'B']

                                base_name = os.path.splitext(filename)[0]
                                safe_sheet_name = sheet_name.replace(" ", "_").replace("/", "_")
                                out_filename = f"{base_name}_{safe_sheet_name}_col{i}.csv"
                                out_path = os.path.join(output_dir, out_filename)
                                extracted_df.to_csv(out_path, index=False)

                                log.append(f"✔ {filename} [{sheet_name}] col {i}-{i+1} → {out_filename}")
                    except Exception as e:
                        log.append(f"✖ {filename} [{sheet_name}] Error: {str(e)}")
            except Exception as e:
                log.append(f"✖ {filename} Failed to read: {str(e)}")

    # ログを表示
    print("\n=== 処理ログ ===")
    for entry in log:
        print(entry)

# 実行用
if __name__ == '__main__':
    extract_ab_columns()
