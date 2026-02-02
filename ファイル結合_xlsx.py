# -*- coding: utf-8 -*-
import sys
import os
from datetime import datetime

import pandas as pd

def merge_xlsx(file_paths):
    """
    ・2つの .xlsx ファイルを結合
    ・2つ目以降は先頭行（インデックス0行）を除外
    ・結合後、V列まで（先頭22列）だけを出力
    ・結合元の .xlsx ファイルを自動削除
    """
    # 入力チェック
    if len(file_paths) < 2:
        print("エラー: 2つ以上の .xlsx ファイルを選択して実行してください。")
        return

    # 最初のファイル：ヘッダー付きで読み込み
    first_fp = file_paths[0]
    print(f"読み込み（ヘッダーあり）: {first_fp}")
    try:
        df_merged = pd.read_excel(first_fp, header=0)
    except Exception as e:
        print(f"ファイル読み込みエラー: {first_fp}\n{e}")
        return

    # 2つ目以降：1行目をスキップして読み込み、ヘッダーを揃えて結合
    for fp in file_paths[1:]:
        print(f"読み込み（ヘッダー除外）: {fp}")
        try:
            df_temp = pd.read_excel(fp, header=None, skiprows=1)
        except Exception as e:
            print(f"ファイル読み込みエラー: {fp}\n{e}")
            return
        df_temp.columns = df_merged.columns
        df_merged = pd.concat([df_merged, df_temp], ignore_index=True)

    # V列まで（先頭22列）のみ抽出
    df_merged = df_merged.iloc[:, :22]

    # 出力ファイル名の準備
    base_dir  = os.path.dirname(first_fp)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    xlsx_name = f"merged_{timestamp}.xlsx"
    xlsx_path = os.path.join(base_dir, xlsx_name)

    # .xlsx 形式で保存
    try:
        df_merged.to_excel(xlsx_path, index=False)
        print(f".xlsx 出力完了: {xlsx_path}")
    except Exception as e:
        print(f".xlsx 出力エラー: {e}")
        return

    # 結合元ファイルを削除
    for fp in file_paths:
        try:
            os.remove(fp)
            print(f"元ファイル削除: {fp}")
        except Exception as e:
            print(f"元ファイル削除エラー: {fp}\n{e}")

def main():
    # 送る で渡されたファイルパスを取得
    file_paths = sys.argv[1:]
    merge_xlsx(file_paths)

if __name__ == "__main__":
    main()
