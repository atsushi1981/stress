import os
import pandas as pd
from openpyxl import Workbook

def create_sample_excels(output_dir='./input'):
    os.makedirs(output_dir, exist_ok=True)

    # サンプルデータの構造
    sample_data = {
        'Header': ['X', 'A', 'B', 'C', 'A', 'B'],
        'Row1':   [1, 2, 3, 4, 5, 6],
        'Row2':   [7, 8, 9, 10, 11, 12],
        'Row3':   [13, 14, 15, 16, 17, 18],
    }

    df = pd.DataFrame(sample_data).T.reset_index(drop=True)

    for file_idx in range(1, 3):
        filename = f'sample{file_idx}.xlsx'
        file_path = os.path.join(output_dir, filename)
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            for sheet_idx in range(1, 3):
                sheet_name = f'Sheet{sheet_idx}'
                df.to_excel(writer, sheet_name=sheet_name, header=False, index=False)

    print("✅ サンプルExcelファイルを生成しました → ./input/sample1.xlsx, sample2.xlsx")

# 実行用
if __name__ == '__main__':
    create_sample_excels()
