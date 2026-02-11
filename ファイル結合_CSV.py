import sys
import os
import csv
import getpass
import shutil
import sqlite3
import pandas as pd
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import time
import threading

def get_database_path():
    """
    Customer_List.dbのパスを取得する
    
    Returns:
        Path: 存在するデータベースファイルパス、見つからない場合はNone
    """
    username = getpass.getuser()
    
    db_paths = [
        f"C:\\Users\\{username}\\OneDrive - 東邦ヤンマーテック株式会社\\CR推進本部フォルダ\\06_社内管理資料\\miraimiru移行関連\\フォルダ共有テスト\\DB\\売上_受注残データ加工用\\Customer_List.db",
        f"C:\\Users\\{username}\\東邦ヤンマーテック株式会社\\CR推進本部フォルダ\\06_社内管理資料\\miraimiru移行関連\\フォルダ共有テスト\\DB\\売上_受注残データ加工用\\Customer_List.db",
        f"C:\\Users\\{username}\\東邦ヤンマーテック株式会社\\CR推進本部 - CR推進本部フォルダ\\06_社内管理資料\\miraimiru移行関連\\フォルダ共有テスト\\DB\\売上_受注残データ加工用\\Customer_List.db"
    ]
    
    for db_path in db_paths:
        if os.path.exists(db_path):
            print(f"データベースが見つかりました: {db_path}")
            return Path(db_path)
    
    print("⚠️ Customer_List.dbが見つかりません")
    print("確認されたパス候補:")
    for i, path in enumerate(db_paths, 1):
        print(f"  {i}. {path}")
    
    return None

def get_order_database_path():
    """
    受注案件状況確認DBの保存パスを取得する
    
    Returns:
        Path: 存在するデータベースフォルダパス、見つからない場合はNone
    """
    username = getpass.getuser()
    
    db_folder_paths = [
        f"C:\\Users\\{username}\\OneDrive - 東邦ヤンマーテック株式会社\\CR推進本部フォルダ\\06_社内管理資料\\miraimiru移行関連\\フォルダ共有テスト\\DB\\受注案件状況確認",
        f"C:\\Users\\{username}\\東邦ヤンマーテック株式会社\\CR推進本部フォルダ\\06_社内管理資料\\miraimiru移行関連\\フォルダ共有テスト\\DB\\受注案件状況確認",
        f"C:\\Users\\{username}\\東邦ヤンマーテック株式会社\\CR推進本部 - CR推進本部フォルダ\\06_社内管理資料\\miraimiru移行関連\\フォルダ共有テスト\\DB\\受注案件状況確認"
    ]
    
    for db_folder_path in db_folder_paths:
        if os.path.exists(db_folder_path):
            print(f"受注案件状況確認DB フォルダが見つかりました: {db_folder_path}")
            return Path(db_folder_path)
    
    print("⚠️ 受注案件状況確認DB フォルダが見つかりません")
    print("確認されたパス候補:")
    for i, path in enumerate(db_folder_paths, 1):
        print(f"  {i}. {path}")
    
    return None

def get_order_data_database_path():
    """
    受注データDBの保存パスを取得する

    Returns:
        Path: 存在するデータベースフォルダパス、見つからない場合はNone
    """
    username = getpass.getuser()

    db_folder_paths = [
        f"C:\\Users\\{username}\\OneDrive - 東邦ヤンマーテック株式会社\\CR推進本部フォルダ\\06_社内管理資料\\miraimiru移行関連\\フォルダ共有テスト\\DB\\受注データ",
        f"C:\\Users\\{username}\\東邦ヤンマーテック株式会社\\CR推進本部フォルダ\\06_社内管理資料\\miraimiru移行関連\\フォルダ共有テスト\\DB\\受注データ",
        f"C:\\Users\\{username}\\東邦ヤンマーテック株式会社\\CR推進本部 - CR推進本部フォルダ\\06_社内管理資料\\miraimiru移行関連\\フォルダ共有テスト\\DB\\受注データ"
    ]

    for db_folder_path in db_folder_paths:
        if os.path.exists(db_folder_path):
            print(f"受注データDB フォルダが見つかりました: {db_folder_path}")
            return Path(db_folder_path)

    # パスが存在しない場合は作成を試行
    for db_folder_path in db_folder_paths:
        try:
            os.makedirs(db_folder_path, exist_ok=True)
            print(f"受注データDBフォルダを作成しました: {db_folder_path}")
            return Path(db_folder_path)
        except Exception as e:
            print(f"フォルダ作成失敗: {db_folder_path} - {str(e)}")
            continue

    print("⚠️ 受注データDB フォルダが見つからず、作成もできませんでした")
    return None

def get_order_info_database_path():
    """
    発注情報DBの保存パスを取得する

    Returns:
        Path: 存在するデータベースフォルダパス、見つからない場合はNone
    """
    username = getpass.getuser()

    db_folder_paths = [
        f"C:\\Users\\{username}\\OneDrive - 東邦ヤンマーテック株式会社\\CR推進本部フォルダ\\06_社内管理資料\\miraimiru移行関連\\フォルダ共有テスト\\現場用\\DB\\発注情報",
        f"C:\\Users\\{username}\\東邦ヤンマーテック株式会社\\CR推進本部フォルダ\\06_社内管理資料\\miraimiru移行関連\\フォルダ共有テスト\\現場用\\DB\\発注情報",
        f"C:\\Users\\{username}\\東邦ヤンマーテック株式会社\\CR推進本部 - CR推進本部フォルダ\\06_社内管理資料\\miraimiru移行関連\\フォルダ共有テスト\\現場用\\DB\\発注情報"
    ]
    
    for db_folder_path in db_folder_paths:
        if os.path.exists(db_folder_path):
            print(f"発注情報DB フォルダが見つかりました: {db_folder_path}")
            return Path(db_folder_path)
    
    # パスが存在しない場合は作成を試行
    for db_folder_path in db_folder_paths:
        try:
            os.makedirs(db_folder_path, exist_ok=True)
            print(f"発注情報DBフォルダを作成しました: {db_folder_path}")
            return Path(db_folder_path)
        except Exception as e:
            print(f"フォルダ作成失敗: {db_folder_path} - {str(e)}")
            continue
    
    print("⚠️ 発注情報DB フォルダが見つからず、作成もできませんでした")
    return None

def get_outstanding_orders_database_path():
    """
    Outstanding_orders.dbの保存パスを取得する
    
    Returns:
        Path: 存在するデータベースフォルダパス、見つからない場合はNone
    """
    username = getpass.getuser()
    
    db_folder_paths = [
        f"C:\\Users\\{username}\\OneDrive - 東邦ヤンマーテック株式会社\\CR推進本部フォルダ\\06_社内管理資料\\miraimiru移行関連\\フォルダ共有テスト\\DB\\発注データ",
        f"C:\\Users\\{username}\\東邦ヤンマーテック株式会社\\CR推進本部フォルダ\\06_社内管理資料\\miraimiru移行関連\\フォルダ共有テスト\\DB\\発注データ",
        f"C:\\Users\\{username}\\東邦ヤンマーテック株式会社\\CR推進本部 - CR推進本部フォルダ\\06_社内管理資料\\miraimiru移行関連\\フォルダ共有テスト\\DB\\発注データ"
    ]
    
    for db_folder_path in db_folder_paths:
        if os.path.exists(db_folder_path):
            print(f"発注データDB フォルダが見つかりました: {db_folder_path}")
            return Path(db_folder_path)
    
    # パスが存在しない場合は作成を試行
    for db_folder_path in db_folder_paths:
        try:
            os.makedirs(db_folder_path, exist_ok=True)
            print(f"発注データDBフォルダを作成しました: {db_folder_path}")
            return Path(db_folder_path)
        except Exception as e:
            print(f"フォルダ作成失敗: {db_folder_path} - {str(e)}")
            continue
    
    print("⚠️ 発注データDB フォルダが見つからず、作成もできませんでした")
    return None

def get_arrival_data_database_path():
    """
    入荷データDBの保存パスを取得する

    Returns:
        Path: 存在するデータベースフォルダパス、見つからない場合はNone
    """
    username = getpass.getuser()

    db_folder_paths = [
        f"C:\\Users\\{username}\\OneDrive - 東邦ヤンマーテック株式会社\\CR推進本部フォルダ\\06_社内管理資料\\miraimiru移行関連\\フォルダ共有テスト\\DB\\入荷データ",
        f"C:\\Users\\{username}\\東邦ヤンマーテック株式会社\\CR推進本部フォルダ\\06_社内管理資料\\miraimiru移行関連\\フォルダ共有テスト\\DB\\入荷データ",
        f"C:\\Users\\{username}\\東邦ヤンマーテック株式会社\\CR推進本部 - CR推進本部フォルダ\\06_社内管理資料\\miraimiru移行関連\\フォルダ共有テスト\\DB\\入荷データ"
    ]

    for db_folder_path in db_folder_paths:
        if os.path.exists(db_folder_path):
            print(f"入荷データDB フォルダが見つかりました: {db_folder_path}")
            return Path(db_folder_path)

    # パスが存在しない場合は作成を試行
    for db_folder_path in db_folder_paths:
        try:
            os.makedirs(db_folder_path, exist_ok=True)
            print(f"入荷データDBフォルダを作成しました: {db_folder_path}")
            return Path(db_folder_path)
        except Exception as e:
            print(f"フォルダ作成失敗: {db_folder_path} - {str(e)}")
            continue

    print("⚠️ 入荷データDB フォルダが見つからず、作成もできませんでした")
    return None

def init_outstanding_orders_database(db_file):
    """Outstanding_ordersデータベースを初期化"""
    try:
        # 既存のDBファイルを削除
        if db_file.exists():
            db_file.unlink()
            print(f"既存のDBファイルを削除しました: {db_file}")
        
        conn = sqlite3.connect(str(db_file), timeout=30)
        cursor = conn.cursor()
        
        # 280列のカラムを作成（column_1 から column_280）
        columns_sql = ", ".join([f"column_{i} TEXT" for i in range(1, 281)])
        
        # テーブルを作成
        cursor.execute(f"""
            CREATE TABLE outstanding_orders (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                {columns_sql},
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        """)
        
        conn.commit()
        conn.close()
        print(f"Outstanding_ordersデータベースを初期化しました: {db_file}")
        
    except Exception as e:
        print(f"Outstanding_ordersデータベース初期化エラー: {str(e)}")
        raise

def create_outstanding_orders_database(csv_file_path):
    """
    Outstanding_orders.csvからOutstanding_orders.dbを作成する
    
    Args:
        csv_file_path: Outstanding_orders.csvのファイルパス
    
    Returns:
        bool: 処理成功時True、失敗時False
    """
    try:
        # データベース保存フォルダのパスを取得
        db_folder = get_outstanding_orders_database_path()
        if db_folder is None:
            print("⚠️ 発注データDB フォルダが見つからないため、DB作成をスキップします")
            return False
        
        print("\n=== Outstanding_orders.db作成処理を開始 ===")
        
        # データベースファイルパス
        db_file_path = db_folder / "Outstanding_orders.db"
        
        # データベース初期化
        init_outstanding_orders_database(db_file_path)
        
        # データベースに接続
        conn = sqlite3.connect(str(db_file_path), timeout=30)
        cursor = conn.cursor()
        
        # CSVファイルを読み込んでデータを挿入
        with open(csv_file_path, 'r', encoding='cp932') as csvfile:
            csv_reader = csv.reader(csvfile)
            
            inserted_count = 0
            error_count = 0
            
            # INSERT文を準備（280列分のプレースホルダー）
            placeholders = ", ".join(["?" for _ in range(280)])
            column_names = ", ".join([f"column_{i}" for i in range(1, 281)])
            insert_sql = f"INSERT INTO outstanding_orders ({column_names}) VALUES ({placeholders})"
            
            for row in csv_reader:
                try:
                    # 行のデータが280列に満たない場合は空文字で埋める
                    while len(row) < 280:
                        row.append("")
                    
                    # 280列を超える場合は280列までに制限
                    if len(row) > 280:
                        row = row[:280]
                    
                    # データを挿入
                    cursor.execute(insert_sql, row)
                    inserted_count += 1
                    
                except Exception as e:
                    error_count += 1
                    print(f"⚠️ データ挿入エラー (行 {inserted_count + error_count}): {str(e)}")
                    continue
        
        # コミットして接続を閉じる
        conn.commit()
        conn.close()
        
        # ファイルサイズを確認
        db_size = db_file_path.stat().st_size
        
        print(f"Outstanding_orders.db作成完了!")
        print(f"保存先: {db_file_path}")
        print(f"挿入レコード数: {inserted_count:,} 件")
        if error_count > 0:
            print(f"エラー件数: {error_count:,} 件")
        print(f"DBファイルサイズ: {db_size:,} bytes ({db_size/1024/1024:.2f} MB)")
        
        return True
        
    except Exception as e:
        print(f"⚠️ Outstanding_orders.db作成処理でエラーが発生しました: {e}")
        return False

def process_order_outstanding_csv(csv_file_path):
    """
    【標準】_発注ファイルから発注残ファイル(Outstanding orders.csv)を作成し、データベースに保存する
    
    Args:
        csv_file_path: 処理対象の【標準】_発注CSVファイルパス
    
    Returns:
        tuple: (Outstanding orders.csv作成成功, Outstanding_orders.db作成成功)
    """
    try:
        print("\n=== 発注残ファイル作成処理を開始 ===")
        
        # 発注実績ファイルを読み込む
        df = pd.read_csv(csv_file_path, encoding='cp932')
        print(f"元ファイル行数: {len(df)} 行")
        
        # 新しいデータフレームを作成（280列）
        new_df = pd.DataFrame(index=df.index, columns=range(280))
        
        # データ抽出と配置（元プログラムのマッピングに基づく）
        new_df[1] = df.iloc[:, 0]     # インデックス1 <- インデックス0
        new_df[49] = df.iloc[:, 53]   # インデックス49 <- インデックス53
        new_df[201] = df.iloc[:, 141] # インデックス201 <- インデックス141
        new_df[205] = df.iloc[:, 145] # インデックス205 <- インデックス145
        new_df[206] = df.iloc[:, 146] # インデックス206 <- インデックス146
        new_df[221] = df.iloc[:, 160] # インデックス221 <- インデックス160
        new_df[2] = df.iloc[:, 1]     # インデックス2 <- インデックス1
        new_df[3] = df.iloc[:, 2]     # インデックス3 <- インデックス2
        new_df[5] = df.iloc[:, 3]     # インデックス5 <- インデックス3
        new_df[13] = df.iloc[:, 10]   # インデックス13 <- インデックス10
        new_df[18] = df.iloc[:, 17]   # インデックス18 <- インデックス17
        new_df[19] = df.iloc[:, 18]   # インデックス19 <- インデックス18
        new_df[23] = df.iloc[:, 23]   # インデックス23 <- インデックス23
        new_df[26] = df.iloc[:, 27]   # インデックス26 <- インデックス27
        new_df[48] = df.iloc[:, 51]   # インデックス48 <- インデックス51
        new_df[50] = df.iloc[:, 54]   # インデックス50 <- インデックス54
        new_df[51] = df.iloc[:, 55]   # インデックス51 <- インデックス55
        new_df[52] = df.iloc[:, 56]   # インデックス52 <- インデックス56
        new_df[118] = df.iloc[:, 87]  # インデックス118 <- インデックス87
        new_df[149] = df.iloc[:, 118] # インデックス149 <- インデックス118
        new_df[150] = df.iloc[:, 119] # インデックス150 <- インデックス119
        new_df[151] = df.iloc[:, 121] # インデックス151 <- インデックス121
        new_df[152] = df.iloc[:, 122] # インデックス152 <- インデックス122
        new_df[155] = df.iloc[:, 210] # インデックス155 <- インデックス210
        new_df[156] = df.iloc[:, 211] # インデックス156 <- インデックス211
        new_df[204] = df.iloc[:, 55]  # インデックス204 <- インデックス55
        new_df[229] = df.iloc[:, 106] # インデックス229 <- インデックス106
        new_df[222] = df.iloc[:, 162] # インデックス222 <- インデックス162
        new_df[244] = df.iloc[:, 252] # インデックス244 <- インデックス252
        
        # 列番号を1から280に設定
        new_df.columns = range(1, 281)
        
        # 同じフォルダに新しいCSVファイルとして保存（ヘッダー行なし）
        folder_path = Path(csv_file_path).parent
        output_file = folder_path / "Outstanding orders.csv"
        new_df.to_csv(output_file, index=False, header=False, encoding='cp932')
        
        print(f"発注残ファイルを作成しました: {output_file}")
        
        # データベースに発注情報を格納
        save_order_info_to_database(df)
        
        # Outstanding_orders.dbを作成
        outstanding_db_success = create_outstanding_orders_database(output_file)
        
        return (True, outstanding_db_success)
        
    except Exception as e:
        print(f"⚠️ 発注残ファイル作成処理でエラーが発生しました: {e}")
        return (False, False)  # ← ここを修正（タプルを返す）

def get_shipping_database_path():
    """
    出荷依頼関連DBの保存パスを取得する
    
    Returns:
        Path: 存在するデータベースフォルダパス、見つからない場合はNone
    """
    username = getpass.getuser()
    
    db_folder_paths = [
        f"C:\\Users\\{username}\\OneDrive - 東邦ヤンマーテック株式会社\\CR推進本部フォルダ\\06_社内管理資料\\miraimiru移行関連\\フォルダ共有テスト\\現場用\\DB\\出荷依頼関連",
        f"C:\\Users\\{username}\\東邦ヤンマーテック株式会社\\CR推進本部フォルダ\\06_社内管理資料\\miraimiru移行関連\\フォルダ共有テスト\\現場用\\DB\\出荷依頼関連",
        f"C:\\Users\\{username}\\東邦ヤンマーテック株式会社\\CR推進本部 - CR推進本部フォルダ\\06_社内管理資料\\miraimiru移行関連\\フォルダ共有テスト\\現場用\\DB\\出荷依頼関連"
    ]
    
    for db_folder_path in db_folder_paths:
        if os.path.exists(db_folder_path):
            print(f"出荷依頼関連DB フォルダが見つかりました: {db_folder_path}")
            return Path(db_folder_path)
    
    # パスが存在しない場合は作成を試行
    for db_folder_path in db_folder_paths:
        try:
            os.makedirs(db_folder_path, exist_ok=True)
            print(f"出荷依頼関連DBフォルダを作成しました: {db_folder_path}")
            return Path(db_folder_path)
        except Exception as e:
            print(f"フォルダ作成失敗: {db_folder_path} - {str(e)}")
            continue
    
    print("⚠️ 出荷依頼関連DB フォルダが見つからず、作成もできませんでした")
    return None

def get_purchase_price_database_path():
    """
    Purchase_price.dbの保存パスを取得する
    
    Returns:
        Path: 存在するデータベースフォルダパス、見つからない場合はNone
    """
    username = getpass.getuser()
    
    db_paths = [
        f"C:\\Users\\{username}\\OneDrive - 東邦ヤンマーテック株式会社\\CR推進本部フォルダ\\06_社内管理資料\\miraimiru移行関連\\フォルダ共有テスト\\DB\\売上_受注残データ加工用\\Purchase_price.db",
        f"C:\\Users\\{username}\\東邦ヤンマーテック株式会社\\CR推進本部フォルダ\\06_社内管理資料\\miraimiru移行関連\\フォルダ共有テスト\\DB\\売上_受注残データ加工用\\Purchase_price.db",
        f"C:\\Users\\{username}\\東邦ヤンマーテック株式会社\\CR推進本部 - CR推進本部フォルダ\\06_社内管理資料\\miraimiru移行関連\\フォルダ共有テスト\\DB\\売上_受注残データ加工用\\Purchase_price.db"
    ]
    
    for db_path in db_paths:
        db_dir = os.path.dirname(db_path)
        if os.path.exists(db_dir):
            print(f"Purchase_price.db 保存フォルダが見つかりました: {db_dir}")
            return Path(db_path)
    
    print("⚠️ Purchase_price.db 保存フォルダが見つかりません")
    print("確認されたパス候補:")
    for i, path in enumerate(db_paths, 1):
        print(f"  {i}. {path}")
    
    return None

def get_order_csv_path():
    """
    【標準】_受注.csvのパスを取得する
    
    Returns:
        Path: 存在するCSVファイルパス、見つからない場合はNone
    """
    username = getpass.getuser()
    
    csv_paths = [
        f"C:\\Users\\{username}\\OneDrive - 東邦ヤンマーテック株式会社\\CR推進本部フォルダ\\06_社内管理資料\\miraimiru移行関連\\フォルダ共有テスト\\受注残\\【標準】_受注.csv",
        f"C:\\Users\\{username}\\東邦ヤンマーテック株式会社\\CR推進本部フォルダ\\06_社内管理資料\\miraimiru移行関連\\フォルダ共有テスト\\受注残\\【標準】_受注.csv",
        f"C:\\Users\\{username}\\東邦ヤンマーテック株式会社\\CR推進本部 - CR推進本部フォルダ\\06_社内管理資料\\miraimiru移行関連\\フォルダ共有テスト\\受注残\\【標準】_受注.csv"
    ]
    
    for csv_path in csv_paths:
        if os.path.exists(csv_path):
            print(f"【標準】_受注.csvが見つかりました: {csv_path}")
            return Path(csv_path)
    
    print("⚠️ 【標準】_受注.csvが見つかりません")
    print("確認されたパス候補:")
    for i, path in enumerate(csv_paths, 1):
        print(f"  {i}. {path}")
    
    return None

def create_order_status_database(csv_file_path):
    """
    受注CSVファイルから受注案件状況確認DBを作成する
    
    Args:
        csv_file_path: 処理対象のCSVファイルパス
    
    Returns:
        bool: 処理成功時True、失敗時False
    """
    try:
        # データベース保存フォルダのパスを取得
        db_folder = get_order_database_path()
        if db_folder is None:
            print("⚠️ 受注案件状況確認DB フォルダが見つからないため、DB作成をスキップします")
            return False
        
        print("\n=== 受注案件状況確認DB作成処理を開始 ===")
        
        # データベースファイルパス
        db_file_path = db_folder / "受注案件状況確認.db"
        
        # 抽出対象のカラム名
        target_columns = [
            "受注番号", "受注件名", "受注日", "部門コード", "部門名", "社員コード", "社員名", 
            "得意先", "得意先名", "受渡場所名", "計上基準区分", "計上基準区分名", "摘要", 
            "最終出荷依頼日", "出荷依頼状況区分", "出荷依頼状況区分名", "最終売上日", 
            "売上状況区分", "売上状況区分名", "更新担当者名", "出荷依頼摘要", "明細_倉庫コード", 
            "明細_商品コード", "明細_商品受注名", "明細_受注数量", "明細_発注引当仕入数量", 
            "明細_自社在庫引当数量", "明細_売上返品数量", "共通項目2", "共通項目2名", 
            "明細_共通項目1", "明細_共通項目2", "明細_共通項目3"
        ]
        
        # CSVファイルを読み込んでヘッダー行を確認
        with open(csv_file_path, 'r', encoding='cp932') as csvfile:
            csv_reader = csv.reader(csvfile)
            header_row = next(csv_reader)
            
            # 対象カラムのインデックスを取得
            column_indices = {}
            missing_columns = []
            
            for col_name in target_columns:
                try:
                    column_indices[col_name] = header_row.index(col_name)
                except ValueError:
                    missing_columns.append(col_name)
            
            if missing_columns:
                print(f"⚠️ 以下のカラムがCSVに見つかりません: {missing_columns}")
                # 見つからないカラムがあっても、見つかったカラムで処理を続行
            
            print(f"抽出対象カラム数: {len(column_indices)}/{len(target_columns)}")
        
        # データベース接続（既存ファイルがある場合は削除して新規作成）
        if db_file_path.exists():
            db_file_path.unlink()
            print(f"既存のDBファイルを削除しました: {db_file_path}")
        
        conn = sqlite3.connect(str(db_file_path), timeout=30)
        cursor = conn.cursor()
        
        # テーブル作成（存在するカラムのみ）
        existing_columns = [col for col in target_columns if col in column_indices]
        columns_sql = ", ".join([f'"{col}" TEXT' for col in existing_columns])
        
        create_table_sql = f'''
        CREATE TABLE 受注案件状況確認 (
            {columns_sql}
        )
        '''
        cursor.execute(create_table_sql)
        print(f"テーブル「受注案件状況確認」を作成しました")
        
        # データ挿入
        insert_sql = f'''
        INSERT INTO 受注案件状況確認 ({", ".join([f'"{col}"' for col in existing_columns])})
        VALUES ({", ".join(["?" for _ in existing_columns])})
        '''
        
        # CSVファイルを再度読み込んでデータを挿入
        with open(csv_file_path, 'r', encoding='cp932') as csvfile:
            csv_reader = csv.reader(csvfile)
            next(csv_reader)  # ヘッダー行をスキップ
            
            inserted_count = 0
            for row in csv_reader:
                # 対象カラムのデータを抽出
                extracted_data = []
                for col_name in existing_columns:
                    col_index = column_indices[col_name]
                    if col_index < len(row):
                        extracted_data.append(row[col_index])
                    else:
                        extracted_data.append("")  # データが不足している場合は空文字
                
                cursor.execute(insert_sql, extracted_data)
                inserted_count += 1
        
        # コミットして接続を閉じる
        conn.commit()
        conn.close()
        
        # ファイルサイズを確認
        db_size = db_file_path.stat().st_size
        
        print(f"受注案件状況確認DB作成完了!")
        print(f"保存先: {db_file_path}")
        print(f"挿入レコード数: {inserted_count:,} 件")
        print(f"DBファイルサイズ: {db_size:,} bytes ({db_size/1024/1024:.2f} MB)")
        
        return True
        
    except Exception as e:
        print(f"⚠️ 受注案件状況確認DB作成処理でエラーが発生しました: {e}")
        return False

def init_order_data_database(db_file, header_columns):
    """受注データデータベースを初期化（CSVヘッダーをカラム名として使用）"""
    try:
        # 既存のDBファイルを削除
        if db_file.exists():
            db_file.unlink()
            print(f"既存のDBファイルを削除しました: {db_file}")

        conn = sqlite3.connect(str(db_file), timeout=30)
        cursor = conn.cursor()

        # CSVヘッダーをカラム名として使用（ダブルクォートで囲んで特殊文字に対応）
        columns_sql = ", ".join([f'"{col}" TEXT' for col in header_columns])

        # テーブルを作成
        cursor.execute(f"""
            CREATE TABLE order_data (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                {columns_sql},
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        """)

        conn.commit()
        conn.close()
        print(f"受注データデータベースを初期化しました: {db_file}")

    except Exception as e:
        print(f"受注データデータベース初期化エラー: {str(e)}")
        raise

def create_order_data_database(csv_file_path):
    """
    受注CSVファイルから受注データDBを作成する（CSVと同内容）

    Args:
        csv_file_path: 処理対象のCSVファイルパス

    Returns:
        bool: 処理成功時True、失敗時False
    """
    try:
        # データベース保存フォルダのパスを取得
        db_folder = get_order_data_database_path()
        if db_folder is None:
            print("⚠️ 受注データDB フォルダが見つからないため、DB作成をスキップします")
            return False

        print("\n=== 受注データDB作成処理を開始 ===")

        # データベースファイルパス
        db_file_path = db_folder / "order_data.db"

        # CSVファイルを読み込んでヘッダー行を取得
        with open(csv_file_path, 'r', encoding='cp932') as csvfile:
            csv_reader = csv.reader(csvfile)
            header_row = next(csv_reader)
            num_columns = len(header_row)
            print(f"CSVカラム数: {num_columns}")

        # データベース初期化（ヘッダー行をカラム名として使用）
        init_order_data_database(db_file_path, header_row)

        # データベースに接続
        conn = sqlite3.connect(str(db_file_path), timeout=30)
        cursor = conn.cursor()

        # CSVファイルを読み込んでデータを挿入（ヘッダー行はスキップ）
        with open(csv_file_path, 'r', encoding='cp932') as csvfile:
            csv_reader = csv.reader(csvfile)
            next(csv_reader)  # ヘッダー行をスキップ

            inserted_count = 0
            error_count = 0

            # INSERT文を準備（カラム名はCSVヘッダーを使用）
            placeholders = ", ".join(["?" for _ in range(num_columns)])
            column_names = ", ".join([f'"{col}"' for col in header_row])
            insert_sql = f"INSERT INTO order_data ({column_names}) VALUES ({placeholders})"

            for row in csv_reader:
                try:
                    # 行のデータがnum_columns列に満たない場合は空文字で埋める
                    while len(row) < num_columns:
                        row.append("")

                    # num_columnsを超える場合はnum_columns列までに制限
                    if len(row) > num_columns:
                        row = row[:num_columns]

                    # データを挿入
                    cursor.execute(insert_sql, row)
                    inserted_count += 1

                except Exception as e:
                    error_count += 1
                    print(f"⚠️ データ挿入エラー (行 {inserted_count + error_count}): {str(e)}")
                    continue

        # コミットして接続を閉じる
        conn.commit()
        conn.close()

        # ファイルサイズを確認
        db_size = db_file_path.stat().st_size

        print(f"受注データDB作成完了!")
        print(f"保存先: {db_file_path}")
        print(f"挿入レコード数: {inserted_count:,} 件")
        if error_count > 0:
            print(f"エラー件数: {error_count:,} 件")
        print(f"DBファイルサイズ: {db_size:,} bytes ({db_size/1024/1024:.2f} MB)")

        return True

    except Exception as e:
        print(f"⚠️ 受注データDB作成処理でエラーが発生しました: {e}")
        return False

def init_purchase_order_data_database(db_file, header_columns):
    """発注データデータベースを初期化（CSVヘッダーをカラム名として使用）"""
    try:
        # 既存のDBファイルを削除
        if db_file.exists():
            db_file.unlink()
            print(f"既存のDBファイルを削除しました: {db_file}")

        conn = sqlite3.connect(str(db_file), timeout=30)
        cursor = conn.cursor()

        # CSVヘッダーをカラム名として使用（ダブルクォートで囲んで特殊文字に対応）
        columns_sql = ", ".join([f'"{col}" TEXT' for col in header_columns])

        # テーブルを作成
        cursor.execute(f"""
            CREATE TABLE purchase_order_data (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                {columns_sql},
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        """)

        conn.commit()
        conn.close()
        print(f"発注データデータベースを初期化しました: {db_file}")

    except Exception as e:
        print(f"発注データデータベース初期化エラー: {str(e)}")
        raise

def create_purchase_order_data_database(csv_file_path):
    """
    発注CSVファイルから発注データDBを作成する（CSVと同内容）

    Args:
        csv_file_path: 処理対象のCSVファイルパス

    Returns:
        bool: 処理成功時True、失敗時False
    """
    try:
        # データベース保存フォルダのパスを取得（Outstanding_orders.dbと同じフォルダ）
        db_folder = get_outstanding_orders_database_path()
        if db_folder is None:
            print("⚠️ 発注データDB フォルダが見つからないため、DB作成をスキップします")
            return False

        print("\n=== 発注データDB作成処理を開始 ===")

        # データベースファイルパス
        db_file_path = db_folder / "purchase_order_data.db"

        # CSVファイルを読み込んでヘッダー行を取得
        with open(csv_file_path, 'r', encoding='cp932') as csvfile:
            csv_reader = csv.reader(csvfile)
            header_row = next(csv_reader)
            num_columns = len(header_row)
            print(f"CSVカラム数: {num_columns}")

        # データベース初期化（ヘッダー行をカラム名として使用）
        init_purchase_order_data_database(db_file_path, header_row)

        # データベースに接続
        conn = sqlite3.connect(str(db_file_path), timeout=30)
        cursor = conn.cursor()

        # CSVファイルを読み込んでデータを挿入（ヘッダー行はスキップ）
        with open(csv_file_path, 'r', encoding='cp932') as csvfile:
            csv_reader = csv.reader(csvfile)
            next(csv_reader)  # ヘッダー行をスキップ

            inserted_count = 0
            error_count = 0

            # INSERT文を準備（カラム名はCSVヘッダーを使用）
            placeholders = ", ".join(["?" for _ in range(num_columns)])
            column_names = ", ".join([f'"{col}"' for col in header_row])
            insert_sql = f"INSERT INTO purchase_order_data ({column_names}) VALUES ({placeholders})"

            for row in csv_reader:
                try:
                    # 行のデータがnum_columns列に満たない場合は空文字で埋める
                    while len(row) < num_columns:
                        row.append("")

                    # num_columnsを超える場合はnum_columns列までに制限
                    if len(row) > num_columns:
                        row = row[:num_columns]

                    # データを挿入
                    cursor.execute(insert_sql, row)
                    inserted_count += 1

                except Exception as e:
                    error_count += 1
                    print(f"⚠️ データ挿入エラー (行 {inserted_count + error_count}): {str(e)}")
                    continue

        # コミットして接続を閉じる
        conn.commit()
        conn.close()

        # ファイルサイズを確認
        db_size = db_file_path.stat().st_size

        print(f"発注データDB作成完了!")
        print(f"保存先: {db_file_path}")
        print(f"挿入レコード数: {inserted_count:,} 件")
        if error_count > 0:
            print(f"エラー件数: {error_count:,} 件")
        print(f"DBファイルサイズ: {db_size:,} bytes ({db_size/1024/1024:.2f} MB)")

        return True

    except Exception as e:
        print(f"⚠️ 発注データDB作成処理でエラーが発生しました: {e}")
        return False

def init_arrival_data_database(db_file, header_columns):
    """入荷データデータベースを初期化（CSVヘッダーをカラム名として使用）"""
    try:
        # 既存のDBファイルを削除
        if db_file.exists():
            db_file.unlink()
            print(f"既存のDBファイルを削除しました: {db_file}")

        conn = sqlite3.connect(str(db_file), timeout=30)
        cursor = conn.cursor()

        # CSVヘッダーをカラム名として使用（ダブルクォートで囲んで特殊文字に対応）
        columns_sql = ", ".join([f'"{col}" TEXT' for col in header_columns])

        # テーブルを作成
        cursor.execute(f"""
            CREATE TABLE arrival_data (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                {columns_sql},
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        """)

        conn.commit()
        conn.close()
        print(f"入荷データデータベースを初期化しました: {db_file}")

    except Exception as e:
        print(f"入荷データデータベース初期化エラー: {str(e)}")
        raise

def create_arrival_data_database(csv_file_path):
    """
    入荷CSVファイルから入荷データDBを作成する（CSVと同内容）

    Args:
        csv_file_path: 処理対象のCSVファイルパス

    Returns:
        bool: 処理成功時True、失敗時False
    """
    try:
        # データベース保存フォルダのパスを取得
        db_folder = get_arrival_data_database_path()
        if db_folder is None:
            print("⚠️ 入荷データDB フォルダが見つからないため、DB作成をスキップします")
            return False

        print("\n=== 入荷データDB作成処理を開始 ===")

        # データベースファイルパス
        db_file_path = db_folder / "arrival_data.db"

        # CSVファイルを読み込んでヘッダー行を取得
        with open(csv_file_path, 'r', encoding='cp932') as csvfile:
            csv_reader = csv.reader(csvfile)
            header_row = next(csv_reader)
            num_columns = len(header_row)
            print(f"CSVカラム数: {num_columns}")

        # データベース初期化（ヘッダー行をカラム名として使用）
        init_arrival_data_database(db_file_path, header_row)

        # データベースに接続
        conn = sqlite3.connect(str(db_file_path), timeout=30)
        cursor = conn.cursor()

        # CSVファイルを読み込んでデータを挿入（ヘッダー行はスキップ）
        with open(csv_file_path, 'r', encoding='cp932') as csvfile:
            csv_reader = csv.reader(csvfile)
            next(csv_reader)  # ヘッダー行をスキップ

            inserted_count = 0
            error_count = 0

            # INSERT文を準備（カラム名はCSVヘッダーを使用）
            placeholders = ", ".join(["?" for _ in range(num_columns)])
            column_names = ", ".join([f'"{col}"' for col in header_row])
            insert_sql = f"INSERT INTO arrival_data ({column_names}) VALUES ({placeholders})"

            for row in csv_reader:
                try:
                    # 行のデータがnum_columns列に満たない場合は空文字で埋める
                    while len(row) < num_columns:
                        row.append("")

                    # num_columnsを超える場合はnum_columns列までに制限
                    if len(row) > num_columns:
                        row = row[:num_columns]

                    # データを挿入
                    cursor.execute(insert_sql, row)
                    inserted_count += 1

                except Exception as e:
                    error_count += 1
                    print(f"⚠️ データ挿入エラー (行 {inserted_count + error_count}): {str(e)}")
                    continue

        # コミットして接続を閉じる
        conn.commit()
        conn.close()

        # ファイルサイズを確認
        db_size = db_file_path.stat().st_size

        print(f"入荷データDB作成完了!")
        print(f"保存先: {db_file_path}")
        print(f"挿入レコード数: {inserted_count:,} 件")
        if error_count > 0:
            print(f"エラー件数: {error_count:,} 件")
        print(f"DBファイルサイズ: {db_size:,} bytes ({db_size/1024/1024:.2f} MB)")

        return True

    except Exception as e:
        print(f"⚠️ 入荷データDB作成処理でエラーが発生しました: {e}")
        return False

def init_shipping_database(db_file):
    """出荷情報データベースを初期化"""
    try:
        # フォルダが存在しない場合は作成
        db_dir = os.path.dirname(db_file)
        os.makedirs(db_dir, exist_ok=True)
        
        conn = sqlite3.connect(str(db_file), timeout=30)
        cursor = conn.cursor()
        
        # 既存のテーブルを削除（テーブル構造変更に対応）
        cursor.execute("DROP TABLE IF EXISTS shipped")
        
        # テーブルを作成（UNIQUE制約を追加して重複を防止）
        cursor.execute("""
            CREATE TABLE shipped (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                shipping_ID TEXT,
                Shipping_data TEXT,
                order_no TEXT,
                shipment_ID TEXT,
                estimate_ID TEXT,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(shipping_ID, Shipping_data, order_no, shipment_ID, estimate_ID)
            )
        """)
        
        conn.commit()
        conn.close()
        print(f"出荷情報データベースを初期化しました: {db_file}")
        
    except Exception as e:
        print(f"出荷情報データベース初期化エラー: {str(e)}")
        raise

def create_shipping_database(csv_file_path):
    """
    出荷CSVファイルから出荷情報DBを作成する
    
    Args:
        csv_file_path: 処理対象のCSVファイルパス
    
    Returns:
        bool: 処理成功時True、失敗時False
    """
    try:
        # データベース保存フォルダのパスを取得
        db_folder = get_shipping_database_path()
        if db_folder is None:
            print("⚠️ 出荷依頼関連DB フォルダが見つからないため、DB作成をスキップします")
            return False
        
        print("\n=== 出荷情報DB作成処理を開始 ===")
        
        # データベースファイルパス
        db_file_path = db_folder / "shipped.db"
        
        # データベース初期化
        init_shipping_database(db_file_path)
        
        # CSVファイルを読み込んでヘッダー行を確認
        with open(csv_file_path, 'r', encoding='cp932') as csvfile:
            csv_reader = csv.reader(csvfile)
            header_row = next(csv_reader)
            
            # 対象カラムのインデックスを取得
            target_columns = {
                '出荷番号': 'shipping_ID',
                '出荷日': 'Shipping_data',
                '受注番号': 'order_no',
                '出荷依頼番号': 'shipment_ID',
                '明細_共通項目3': 'estimate_ID'
            }
            
            column_indices = {}
            missing_columns = []
            
            for csv_col_name, db_col_name in target_columns.items():
                try:
                    column_indices[db_col_name] = header_row.index(csv_col_name)
                except ValueError:
                    missing_columns.append(csv_col_name)
            
            if missing_columns:
                print(f"⚠️ 以下のカラムがCSVに見つかりません: {missing_columns}")
                if len(missing_columns) == len(target_columns):
                    print("⚠️ 必要なカラムが見つからないため、DB作成を中止します")
                    return False
            
            print(f"抽出対象カラム数: {len(column_indices)}/{len(target_columns)}")
        
        # データベースに接続
        conn = sqlite3.connect(str(db_file_path), timeout=30)
        cursor = conn.cursor()
        
        # CSVファイルを再度読み込んでデータを挿入
        with open(csv_file_path, 'r', encoding='cp932') as csvfile:
            csv_reader = csv.reader(csvfile)
            next(csv_reader)  # ヘッダー行をスキップ
            
            inserted_count = 0
            duplicate_count = 0
            total_count = 0
            
            for row in csv_reader:
                total_count += 1
                
                # 対象カラムのデータを抽出
                shipping_ID = row[column_indices['shipping_ID']] if 'shipping_ID' in column_indices and column_indices['shipping_ID'] < len(row) else ""
                Shipping_data = row[column_indices['Shipping_data']] if 'Shipping_data' in column_indices and column_indices['Shipping_data'] < len(row) else ""
                order_no = row[column_indices['order_no']] if 'order_no' in column_indices and column_indices['order_no'] < len(row) else ""
                shipment_ID = row[column_indices['shipment_ID']] if 'shipment_ID' in column_indices and column_indices['shipment_ID'] < len(row) else ""
                estimate_ID = row[column_indices['estimate_ID']] if 'estimate_ID' in column_indices and column_indices['estimate_ID'] < len(row) else ""
                
                try:
                    # INSERT OR IGNOREで重複を無視
                    cursor.execute("""
                        INSERT OR IGNORE INTO shipped (
                            shipping_ID, Shipping_data, order_no, shipment_ID, estimate_ID
                        ) VALUES (?, ?, ?, ?, ?)
                    """, (shipping_ID, Shipping_data, order_no, shipment_ID, estimate_ID))
                    
                    # 実際に挿入されたかチェック（rowcountが0なら重複）
                    if cursor.rowcount > 0:
                        inserted_count += 1
                    else:
                        duplicate_count += 1
                
                except sqlite3.IntegrityError:
                    # UNIQUE制約違反（重複データ）
                    duplicate_count += 1
                    continue
        
        # コミットして接続を閉じる
        conn.commit()
        conn.close()
        
        # ファイルサイズを確認
        db_size = db_file_path.stat().st_size
        
        print(f"出荷情報DB作成完了!")
        print(f"保存先: {db_file_path}")
        print(f"CSVファイル総行数: {total_count:,} 行")
        print(f"挿入レコード数: {inserted_count:,} 件（ユニーク）")
        if duplicate_count > 0:
            print(f"重複除外数: {duplicate_count:,} 件")
        print(f"DBファイルサイズ: {db_size:,} bytes ({db_size/1024/1024:.2f} MB)")
        
        return True
        
    except Exception as e:
        print(f"⚠️ 出荷情報DB作成処理でエラーが発生しました: {e}")
        return False

def init_purchase_price_database(db_file):
    """仕入単価データベースを初期化"""
    try:
        # 既存のDBファイルを削除
        if db_file.exists():
            db_file.unlink()
            print(f"既存のDBファイルを削除しました: {db_file}")
        
        conn = sqlite3.connect(str(db_file), timeout=30)
        cursor = conn.cursor()
        
        # テーブルを作成
        cursor.execute("""
            CREATE TABLE purchase_prices (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                estimate_ID TEXT,
                detail_common2 TEXT,
                purchase_price INTEGER,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        """)
        
        conn.commit()
        conn.close()
        print(f"仕入単価データベースを初期化しました: {db_file}")
        
    except Exception as e:
        print(f"仕入単価データベース初期化エラー: {str(e)}")
        raise

def create_purchase_price_database(csv_file_path):
    """
    仕入CSVファイルから仕入単価DBを作成する（赤伝相殺処理付き）
    
    Args:
        csv_file_path: 処理対象のCSVファイルパス
    
    Returns:
        bool: 処理成功時True、失敗時False
    """
    try:
        # データベースファイルパスを取得
        db_file_path = get_purchase_price_database_path()
        if db_file_path is None:
            print("⚠️ Purchase_price.db 保存フォルダが見つからないため、DB作成をスキップします")
            return False
        
        print("\n=== 仕入単価DB作成処理を開始 ===")
        
        # データベース初期化
        init_purchase_price_database(db_file_path)
        
        # CSVファイルを読み込んでヘッダー行を確認
        with open(csv_file_path, 'r', encoding='cp932') as csvfile:
            csv_reader = csv.reader(csvfile)
            header_row = next(csv_reader)
            
            # 対象カラムのインデックスを取得
            target_columns = {
                '明細_共通項目3': 'estimate_ID',
                '明細_共通項目2': 'detail_common2',
                '明細_仕入単価': 'purchase_price',
                '赤伝フラグ': 'red_flag'
            }
            
            column_indices = {}
            missing_columns = []
            
            for csv_col_name, db_col_name in target_columns.items():
                try:
                    column_indices[db_col_name] = header_row.index(csv_col_name)
                except ValueError:
                    missing_columns.append(csv_col_name)
            
            if missing_columns:
                print(f"⚠️ 以下のカラムがCSVに見つかりません: {missing_columns}")
                if '赤伝フラグ' in missing_columns:
                    print("⚠️ 赤伝フラグ列が見つからないため、赤伝相殺処理なしで続行します")
                # estimate_ID, detail_common2, purchase_priceが全て見つからない場合のみ中止
                essential_columns = ['estimate_ID', 'detail_common2', 'purchase_price']
                if all(col not in column_indices for col in essential_columns):
                    print("⚠️ 必要なカラムが見つからないため、DB作成を中止します")
                    return False
            
            print(f"抽出対象カラム数: {len(column_indices)}/{len(target_columns)}")
        
        # 第1パス：赤伝フラグが"1"のレコードの組み合わせを収集
        red_flag_combinations = set()
        red_flag_available = 'red_flag' in column_indices
        
        if red_flag_available:
            print("\n--- 赤伝レコードの特定中 ---")
            with open(csv_file_path, 'r', encoding='cp932') as csvfile:
                csv_reader = csv.reader(csvfile)
                next(csv_reader)  # ヘッダー行をスキップ
                
                red_flag_count = 0
                
                for row in csv_reader:
                    # 赤伝フラグをチェック
                    red_flag = ""
                    if 'red_flag' in column_indices and column_indices['red_flag'] < len(row):
                        red_flag = row[column_indices['red_flag']].strip()
                    
                    if red_flag == "1":
                        # 赤伝レコードの組み合わせを保存
                        estimate_ID = row[column_indices['estimate_ID']] if 'estimate_ID' in column_indices and column_indices['estimate_ID'] < len(row) else ""
                        detail_common2 = row[column_indices['detail_common2']] if 'detail_common2' in column_indices and column_indices['detail_common2'] < len(row) else ""
                        
                        purchase_price = None
                        if 'purchase_price' in column_indices and column_indices['purchase_price'] < len(row):
                            price_str = row[column_indices['purchase_price']].strip()
                            if price_str:
                                try:
                                    purchase_price = int(float(price_str))
                                except ValueError:
                                    continue
                        
                        # 組み合わせをタプルとしてセットに追加
                        combination = (estimate_ID, detail_common2, purchase_price)
                        red_flag_combinations.add(combination)
                        red_flag_count += 1
            
            print(f"赤伝レコード検出数: {red_flag_count} 件")
            print(f"赤伝対象の一意な組み合わせ数: {len(red_flag_combinations)} 件")
        else:
            print("\n⚠️ 赤伝フラグ列がないため、赤伝相殺処理をスキップします")
        
        # 第2パス：データベースへの挿入（赤伝対象を除外）
        print("\n--- データベースへの挿入開始 ---")
        conn = sqlite3.connect(str(db_file_path), timeout=30)
        cursor = conn.cursor()
        
        with open(csv_file_path, 'r', encoding='cp932') as csvfile:
            csv_reader = csv.reader(csvfile)
            next(csv_reader)  # ヘッダー行をスキップ
            
            inserted_count = 0
            error_count = 0
            red_excluded_count = 0
            total_count = 0
            
            for row in csv_reader:
                total_count += 1
                
                # 対象カラムのデータを抽出
                estimate_ID = row[column_indices['estimate_ID']] if 'estimate_ID' in column_indices and column_indices['estimate_ID'] < len(row) else ""
                detail_common2 = row[column_indices['detail_common2']] if 'detail_common2' in column_indices and column_indices['detail_common2'] < len(row) else ""
                
                # 仕入単価を整数値に変換
                purchase_price = None
                if 'purchase_price' in column_indices and column_indices['purchase_price'] < len(row):
                    price_str = row[column_indices['purchase_price']].strip()
                    if price_str:
                        try:
                            # 小数点以下を切り捨てて整数化
                            purchase_price = int(float(price_str))
                        except ValueError:
                            error_count += 1
                            continue
                
                # 赤伝相殺チェック
                if red_flag_available:
                    combination = (estimate_ID, detail_common2, purchase_price)
                    if combination in red_flag_combinations:
                        # 赤伝対象の組み合わせなので除外
                        red_excluded_count += 1
                        continue
                
                try:
                    cursor.execute("""
                        INSERT INTO purchase_prices (
                            estimate_ID, detail_common2, purchase_price
                        ) VALUES (?, ?, ?)
                    """, (estimate_ID, detail_common2, purchase_price))
                    
                    inserted_count += 1
                
                except sqlite3.Error as e:
                    error_count += 1
                    continue
        
        # コミットして接続を閉じる
        conn.commit()
        conn.close()
        
        # ファイルサイズを確認
        db_size = db_file_path.stat().st_size
        
        print(f"\n仕入単価DB作成完了!")
        print(f"保存先: {db_file_path}")
        print(f"CSVファイル総行数: {total_count:,} 行")
        print(f"挿入レコード数: {inserted_count:,} 件")
        if red_flag_available:
            print(f"赤伝相殺除外数: {red_excluded_count:,} 件")
        if error_count > 0:
            print(f"エラー件数: {error_count:,} 件")
        print(f"DBファイルサイズ: {db_size:,} bytes ({db_size/1024/1024:.2f} MB)")
        
        return True
        
    except Exception as e:
        print(f"⚠️ 仕入単価DB作成処理でエラーが発生しました: {e}")
        return False

def load_order_data_from_csv():
    """
    【標準】_受注.csvから受注番号をキーとした辞書を作成
    
    Returns:
        dict: {受注番号: {'受注件名': '', '得意先名': ''}} の辞書、失敗時は空辞書
    """
    try:
        # 受注CSVファイルのパスを取得
        order_csv_path = get_order_csv_path()
        if order_csv_path is None:
            print("⚠️ 受注CSVファイルが見つからないため、受注情報の取得をスキップします")
            return {}
        
        order_data = {}
        
        with open(order_csv_path, 'r', encoding='cp932') as csvfile:
            csv_reader = csv.reader(csvfile)
            header_row = next(csv_reader)
            
            # 必要なカラムのインデックスを取得
            try:
                受注番号_index = header_row.index('受注番号')
                受注件名_index = header_row.index('受注件名')
                得意先名_index = header_row.index('得意先名')
            except ValueError as e:
                print(f"⚠️ 受注CSVに必要なカラムが見つかりません: {e}")
                return {}
            
            # データを読み込んで辞書を作成
            for row in csv_reader:
                if len(row) > max(受注番号_index, 受注件名_index, 得意先名_index):
                    受注番号 = row[受注番号_index].strip()
                    受注件名 = row[受注件名_index].strip()
                    得意先名 = row[得意先名_index].strip()
                    
                    if 受注番号:  # 受注番号が空でない場合のみ
                        order_data[受注番号] = {
                            '受注件名': 受注件名,
                            '得意先名': 得意先名
                        }
        
        print(f"受注CSVから {len(order_data)} 件のデータを読み込みました")
        return order_data
        
    except Exception as e:
        print(f"⚠️ 受注CSVデータ読み込みエラー: {str(e)}")
        return {}

def save_order_info_to_database(df):
    """データベースに発注情報を保存"""
    try:
        db_folder = get_order_info_database_path()
        if not db_folder:
            raise Exception("発注情報DBフォルダが見つかりません")
        
        # データベースファイルのパス
        db_file = db_folder / "orders_info.db"
        
        # データベース初期化
        init_order_info_database(db_file)
        
        # 受注CSVからデータを読み込む
        order_data = load_order_data_from_csv()
        
        # 発注CSVのヘッダー行から「社員名」カラムのインデックスを取得
        社員名_index = None
        if '社員名' in df.columns:
            社員名_index = df.columns.get_loc('社員名')
            print(f"「社員名」カラムが見つかりました（インデックス: {社員名_index}）")
        else:
            print("⚠️ 「社員名」カラムが見つかりません")
        
        # 発注CSVのヘッダー行から「摘要」カラムのインデックスを取得
        摘要_index = None
        if '摘要' in df.columns:
            摘要_index = df.columns.get_loc('摘要')
            print(f"「摘要」カラムが見つかりました（インデックス: {摘要_index}）")
        else:
            print("⚠️ 「摘要」カラムが見つかりません")
        
        # データベースに接続
        conn = sqlite3.connect(str(db_file), timeout=30)
        cursor = conn.cursor()
        
        # 既存データを削除（全件入れ替え）
        cursor.execute("DELETE FROM orders_info")
        
        # フィルタリング条件を適用してデータを挿入
        inserted_count = 0
        skipped_count = 0
        
        for index, row in df.iterrows():
            # 正確な列インデックスに基づいてデータを取得
            発注番号 = row.iloc[0] if len(row) > 0 else ""
            受注番号 = row.iloc[3] if len(row) > 3 else ""
            仕入先 = row.iloc[25] if len(row) > 25 else ""
            仕入先名 = row.iloc[27] if len(row) > 27 else ""
            仕入先略名 = row.iloc[28] if len(row) > 28 else ""
            納期日 = row.iloc[87] if len(row) > 87 else ""
            納入区分名 = row.iloc[52] if len(row) > 52 else ""
            明細_発注数量 = row.iloc[160] if len(row) > 160 else ""
            明細_明細金額 = row.iloc[176] if len(row) > 176 else ""
            明細_倉庫名 = row.iloc[144] if len(row) > 144 else ""
            明細_共通項目2 = row.iloc[250] if len(row) > 250 else ""
            明細_共通項目3 = row.iloc[252] if len(row) > 252 else ""
            明細_商品コード = row.iloc[145] if len(row) > 145 else ""
            明細_商品発注名 = row.iloc[146] if len(row) > 146 else ""
            明細_発注単価 = row.iloc[162] if len(row) > 162 else ""
            共通項目2 = row.iloc[210] if len(row) > 210 else ""
            共通項目2名 = row.iloc[211] if len(row) > 211 else ""
            
            # 社員名を取得
            社員名 = ""
            if 社員名_index is not None and len(row) > 社員名_index:
                社員名 = row.iloc[社員名_index] if row.iloc[社員名_index] is not None else ""
            
            # 摘要を取得
            摘要 = ""
            if 摘要_index is not None and len(row) > 摘要_index:
                摘要 = row.iloc[摘要_index] if row.iloc[摘要_index] is not None else ""
            
            # 受注CSVから受注件名と得意先名を取得
            受注件名 = ""
            得意先名 = ""
            受注番号_str = str(受注番号).strip()
            if 受注番号_str in order_data:
                受注件名 = order_data[受注番号_str]['受注件名']
                得意先名 = order_data[受注番号_str]['得意先名']
            
            # フィルタリング条件をチェック
            if 納入区分名 == "直送" or 明細_倉庫名 == "本社特別在庫":
                skipped_count += 1
                continue
            
            # データを挿入
            cursor.execute("""
                INSERT INTO orders_info (
                    発注番号, 受注番号, 仕入先, 仕入先名, 仕入先略名, 納期日,
                    納入区分名, 明細_発注数量, 明細_明細金額, 明細_倉庫名, 
                    明細_共通項目2, 明細_共通項目3, 明細_商品コード, 明細_商品発注名, 明細_発注単価,
                    共通項目2, 共通項目2名, 受注件名, 得意先名, 社員名, 摘要, 処理日時
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                発注番号,
                受注番号,
                仕入先,
                仕入先名,
                仕入先略名,
                納期日,
                納入区分名,
                明細_発注数量,
                明細_明細金額,
                明細_倉庫名,
                明細_共通項目2,
                明細_共通項目3,
                明細_商品コード,
                明細_商品発注名,
                明細_発注単価,
                共通項目2,
                共通項目2名,
                受注件名,
                得意先名,
                社員名,
                摘要,
                datetime.now().strftime("%Y-%m-%d %H:%M:%S")  # 処理日時
            ))
            inserted_count += 1
        
        conn.commit()
        conn.close()
        
        print(f"発注情報データベースに {inserted_count} 件のレコードを保存しました（{skipped_count} 件をフィルタリングで除外）: {db_file}")
        
    except Exception as e:
        print(f"発注情報データベース保存エラー: {str(e)}")
        raise

def init_order_info_database(db_file):
    """発注情報データベースを初期化"""
    try:
        # フォルダが存在しない場合は作成
        db_dir = os.path.dirname(db_file)
        os.makedirs(db_dir, exist_ok=True)
        
        conn = sqlite3.connect(str(db_file), timeout=30)
        cursor = conn.cursor()
        
        # 既存のテーブルを削除（テーブル構造変更に対応）
        cursor.execute("DROP TABLE IF EXISTS orders_info")
        
        # テーブルを作成
        cursor.execute("""
            CREATE TABLE orders_info (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                発注番号 TEXT,
                受注番号 TEXT,
                仕入先 TEXT,
                仕入先名 TEXT,
                仕入先略名 TEXT,
                納期日 TEXT,
                納入区分名 TEXT,
                明細_発注数量 TEXT,
                明細_明細金額 TEXT,
                明細_倉庫名 TEXT,
                明細_共通項目2 TEXT,
                明細_共通項目3 TEXT,
                明細_商品コード TEXT,
                明細_商品発注名 TEXT,
                明細_発注単価 TEXT,
                共通項目2 TEXT,
                共通項目2名 TEXT,
                受注件名 TEXT,
                得意先名 TEXT,
                社員名 TEXT,
                摘要 TEXT,
                処理日時 TEXT,
                作成日時 DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        """)
        
        conn.commit()
        conn.close()
        print(f"発注情報データベースを初期化しました: {db_file}")
        
    except Exception as e:
        print(f"発注情報データベース初期化エラー: {str(e)}")
        raise

def update_csv_with_employee_data(csv_file_path):
    """
    【標準】_受注のCSVファイルに社員情報を追加する
    
    Args:
        csv_file_path: 処理対象のCSVファイルパス
    
    Returns:
        bool: 処理成功時True、失敗時False
    """
    try:
        # データベースファイルのパスを取得
        db_path = get_database_path()
        if db_path is None:
            print("⚠️ データベースファイルが見つからないため、社員情報の追加をスキップします")
            return False
        
        print("\n=== 社員情報追加処理を開始 ===")
        
        # データベースに接続
        conn = sqlite3.connect(str(db_path), timeout=30)
        cursor = conn.cursor()
        
        # Customer_listテーブルの存在確認（正しいテーブル名で確認）
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='Customer_list'")
        if not cursor.fetchone():
            # Customer_listが見つからない場合は、大文字小文字を無視して検索
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
            tables = [table[0] for table in cursor.fetchall()]
            print(f"⚠️ Customer_listテーブルが見つかりません")
            print(f"利用可能なテーブル: {tables}")
            conn.close()
            return False
        
        # テーブル構造を確認
        cursor.execute("PRAGMA table_info(Customer_list)")
        columns = [column[1] for column in cursor.fetchall()]
        print(f"データベースの列: {columns}")
        
        if 'Customer_ID' not in columns or 'employee_ID' not in columns or 'employee_name' not in columns:
            print("⚠️ 必要な列（Customer_ID, employee_ID, employee_name）が見つかりません")
            conn.close()
            return False
        
        # Customer_IDをキーとした辞書を作成
        cursor.execute("SELECT Customer_ID, employee_ID, employee_name FROM Customer_list")
        customer_data = {}
        for row in cursor.fetchall():
            customer_id, employee_id, employee_name = row
            customer_data[str(customer_id)] = {
                'employee_ID': employee_id if employee_id is not None else '',
                'employee_name': employee_name if employee_name is not None else ''
            }
        
        conn.close()
        print(f"データベースから {len(customer_data)} 件の顧客情報を読み込みました")
        
        # CSVファイルを読み込んで処理
        temp_file = csv_file_path.parent / "temp_employee_update.csv"
        
        with open(csv_file_path, 'r', encoding='cp932') as infile, \
             open(temp_file, 'w', encoding='cp932', newline='') as outfile:
            
            csv_reader = csv.reader(infile)
            csv_writer = csv.writer(outfile, quoting=csv.QUOTE_MINIMAL, lineterminator='\n')
            
            updated_count = 0
            not_found_count = 0
            total_rows = 0
            customer_col_index = None
            not_found_customers = set()
            
            for row_num, row in enumerate(csv_reader):
                total_rows += 1
                
                # ヘッダー行で「得意先」列のインデックスを特定
                if row_num == 0:
                    try:
                        customer_col_index = row.index('得意先')
                        print(f"「得意先」列が見つかりました（インデックス: {customer_col_index}）")
                    except ValueError:
                        print("⚠️ 「得意先」列が見つかりません")
                        csv_writer.writerow(row)
                        continue
                    
                    # ヘッダー行をそのまま出力
                    csv_writer.writerow(row)
                    continue
                
                # データ行の処理
                if customer_col_index is not None and len(row) > max(customer_col_index, 19):
                    customer_id = str(row[customer_col_index]).strip()
                    
                    # データベースから社員情報を取得
                    if customer_id in customer_data:
                        # 社員コード（インデックス18）と社員名（インデックス19）を設定
                        if len(row) > 18:
                            row[18] = customer_data[customer_id]['employee_ID']
                        if len(row) > 19:
                            row[19] = customer_data[customer_id]['employee_name']
                        updated_count += 1
                    else:
                        # データベースに存在しない場合
                        if customer_id:  # 空文字でない場合のみ
                            not_found_customers.add(customer_id)
                            not_found_count += 1
                
                csv_writer.writerow(row)
        
        # 元ファイルを置き換え
        shutil.move(str(temp_file), str(csv_file_path))
        
        print(f"社員情報追加完了: {updated_count}/{total_rows-1} 行を更新")
        if not_found_count > 0:
            print(f"⚠️ データベースに見つからない得意先: {not_found_count} 件")
        
        return True
        
    except Exception as e:
        print(f"⚠️ 社員情報追加処理でエラーが発生しました: {e}")
        # 一時ファイルが存在する場合は削除
        temp_file = csv_file_path.parent / "temp_employee_update.csv"
        if temp_file.exists():
            temp_file.unlink()
        return False

def get_output_folder(selected_filename):
    """
    選択されたファイル名に基づいて出力フォルダパスを取得する
    
    Args:
        selected_filename: 選択されたファイル名
    
    Returns:
        Path: 存在する出力フォルダパス、見つからない場合はNone
    """
    # 現在のユーザー名を取得
    username = getpass.getuser()
    
    # 各ファイル種類毎のパス候補を定義
    folder_paths = {
        "【標準】_受注": [
            f"C:\\Users\\{username}\\OneDrive - 東邦ヤンマーテック株式会社\\CR推進本部フォルダ\\06_社内管理資料\\miraimiru移行関連\\フォルダ共有テスト\\受注残",
            f"C:\\Users\\{username}\\東邦ヤンマーテック株式会社\\CR推進本部フォルダ\\06_社内管理資料\\miraimiru移行関連\\フォルダ共有テスト\\受注残",
            f"C:\\Users\\{username}\\東邦ヤンマーテック株式会社\\CR推進本部 - CR推進本部フォルダ\\06_社内管理資料\\miraimiru移行関連\\フォルダ共有テスト\\受注残"
        ],
        "【標準】_発注": [
            f"C:\\Users\\{username}\\OneDrive - 東邦ヤンマーテック株式会社\\CR推進本部フォルダ\\06_社内管理資料\\miraimiru移行関連\\フォルダ共有テスト\\発注残",
            f"C:\\Users\\{username}\\東邦ヤンマーテック株式会社\\CR推進本部フォルダ\\06_社内管理資料\\miraimiru移行関連\\フォルダ共有テスト\\発注残",
            f"C:\\Users\\{username}\\東邦ヤンマーテック株式会社\\CR推進本部 - CR推進本部フォルダ\\06_社内管理資料\\miraimiru移行関連\\フォルダ共有テスト\\発注残"
        ],
        "【標準】_入荷": [
            f"C:\\Users\\{username}\\OneDrive - 東邦ヤンマーテック株式会社\\CR推進本部フォルダ\\06_社内管理資料\\miraimiru移行関連\\フォルダ共有テスト\\入荷実績",
            f"C:\\Users\\{username}\\東邦ヤンマーテック株式会社\\CR推進本部フォルダ\\06_社内管理資料\\miraimiru移行関連\\フォルダ共有テスト\\入荷実績",
            f"C:\\Users\\{username}\\東邦ヤンマーテック株式会社\\CR推進本部 - CR推進本部フォルダ\\06_社内管理資料\\miraimiru移行関連\\フォルダ共有テスト\\入荷実績"
        ],
        "【標準】_仕入": [
            f"C:\\Users\\{username}\\OneDrive - 東邦ヤンマーテック株式会社\\CR推進本部フォルダ\\06_社内管理資料\\miraimiru移行関連\\フォルダ共有テスト\\仕入実績",
            f"C:\\Users\\{username}\\東邦ヤンマーテック株式会社\\CR推進本部フォルダ\\06_社内管理資料\\miraimiru移行関連\\フォルダ共有テスト\\仕入実績",
            f"C:\\Users\\{username}\\東邦ヤンマーテック株式会社\\CR推進本部 - CR推進本部フォルダ\\06_社内管理資料\\miraimiru移行関連\\フォルダ共有テスト\\仕入実績"
        ],
        "【標準】_出荷": [
            f"C:\\Users\\{username}\\OneDrive - 東邦ヤンマーテック株式会社\\CR推進本部フォルダ\\06_社内管理資料\\miraimiru移行関連\\フォルダ共有テスト\\出荷実績",
            f"C:\\Users\\{username}\\東邦ヤンマーテック株式会社\\CR推進本部フォルダ\\06_社内管理資料\\miraimiru移行関連\\フォルダ共有テスト\\出荷実績",
            f"C:\\Users\\{username}\\東邦ヤンマーテック株式会社\\CR推進本部 - CR推進本部フォルダ\\06_社内管理資料\\miraimiru移行関連\\フォルダ共有テスト\\出荷実績"
        ]
    }
    
    # 選択されたファイル名に対応するパス候補を取得
    if selected_filename not in folder_paths:
        return None
    
    # 存在するパスを検索
    for folder_path in folder_paths[selected_filename]:
        if os.path.exists(folder_path):
            print(f"出力フォルダが見つかりました: {folder_path}")
            return Path(folder_path)
    
    print(f"⚠️ 出力フォルダが見つかりません: {selected_filename}")
    print("確認されたパス候補:")
    for i, path in enumerate(folder_paths[selected_filename], 1):
        print(f"  {i}. {path}")
    
    return None

def show_filename_dialog():
    """
    tkinterを使用したファイル名選択ダイアログを表示
    
    Returns:
        str: 選択されたファイル名（拡張子なし）、キャンセル時はNone
    """
    # メインウィンドウ
    root = tk.Tk()
    root.title("出力ファイル名の選択")
    root.geometry("350x320")
    root.resizable(False, False)
    
    # ウィンドウを画面中央に配置
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (350 // 2)
    y = (root.winfo_screenheight() // 2) - (320 // 2)
    root.geometry(f"350x320+{x}+{y}")
    
    # アイコンを設定（オプション）
    try:
        root.iconbitmap(default='')  # デフォルトアイコン
    except:
        pass
    
    selected_filename = None
    
    # メインフレーム
    main_frame = ttk.Frame(root, padding="20")
    main_frame.pack(fill=tk.BOTH, expand=True)
    
    # タイトルラベル
    title_label = ttk.Label(
        main_frame, 
        text="出力ファイル名を選択してください", 
        font=("", 12, "bold")
    )
    title_label.pack(pady=(0, 20))
    
    # 選択肢
    options = [
        "【標準】_受注",
        "【標準】_発注", 
        "【標準】_入荷",
        "【標準】_仕入",
        "【標準】_出荷"
    ]
    
    # ラジオボタン用の変数
    selected_option = tk.StringVar(value=options[0])
    
    # ラジオボタンフレーム
    radio_frame = ttk.LabelFrame(main_frame, text="ファイル名", padding="15")
    radio_frame.pack(fill=tk.X, pady=(0, 20))
    
    # ラジオボタンを作成
    for option in options:
        radio = ttk.Radiobutton(
            radio_frame,
            text=option,
            variable=selected_option,
            value=option,
            style="TRadiobutton"
        )
        radio.pack(anchor=tk.W, pady=2)
    
    # ボタンフレーム
    button_frame = ttk.Frame(main_frame)
    button_frame.pack(fill=tk.X, pady=(10, 0))
    
    def on_ok():
        nonlocal selected_filename
        selected_filename = selected_option.get()
        root.quit()
        root.destroy()
    
    def on_cancel():
        nonlocal selected_filename
        selected_filename = None
        root.quit()
        root.destroy()
    
    # OKボタン
    ok_button = ttk.Button(
        button_frame,
        text="OK",
        command=on_ok,
        style="Accent.TButton"
    )
    ok_button.pack(side=tk.RIGHT, padx=(5, 0))
    
    # キャンセルボタン
    cancel_button = ttk.Button(
        button_frame,
        text="キャンセル",
        command=on_cancel
    )
    cancel_button.pack(side=tk.RIGHT)
    
    # Enterキーでも実行できるように
    root.bind('<Return>', lambda e: on_ok())
    root.bind('<Escape>', lambda e: on_cancel())
    
    # 最初のラジオボタンにフォーカス
    root.focus_set()
    
    # ダイアログを表示
    root.mainloop()
    
    return selected_filename

def delete_source_files(file_paths):
    """
    結合元のCSVファイルを削除する（削除対象外ファイルを除く）
    
    Args:
        file_paths: 削除対象のファイルパスのリスト
    
    Returns:
        tuple: (削除成功数, 削除スキップ数, 削除失敗数)
    """
    # 削除対象外のファイル名（拡張子含む）
    exclude_files = {
        "【標準】_受注.csv",
        "【標準】_発注.csv",
        "【標準】_入荷.csv",
        "【標準】_仕入.csv",
        "【標準】_出荷.csv"
    }
    
    deleted_count = 0
    skipped_count = 0
    failed_count = 0
    
    print("\n=== 元ファイル削除処理を開始 ===")
    
    for file_path in file_paths:
        file_name = os.path.basename(file_path)
        
        # 削除対象外ファイルかチェック
        if file_name in exclude_files:
            print(f"⏭️  スキップ（削除対象外）: {file_name}")
            skipped_count += 1
            continue
        
        # ファイルが存在するか確認
        if not os.path.exists(file_path):
            print(f"⚠️ ファイルが見つかりません: {file_name}")
            failed_count += 1
            continue
        
        # ファイル削除を試行
        try:
            os.remove(file_path)
            print(f"🗑️  削除完了: {file_name}")
            deleted_count += 1
        except Exception as e:
            print(f"⚠️ 削除失敗: {file_name} - {str(e)}")
            failed_count += 1
    
    # 削除結果サマリー
    print(f"\n削除処理完了:")
    print(f"  ✅ 削除成功: {deleted_count} 件")
    if skipped_count > 0:
        print(f"  ⏭️  削除スキップ: {skipped_count} 件")
    if failed_count > 0:
        print(f"  ⚠️ 削除失敗: {failed_count} 件")
    
    return (deleted_count, skipped_count, failed_count)

def merge_csv_files(file_paths):
    """
    複数のCSVファイルを結合する（クオート除去処理付き）
    
    Args:
        file_paths: 結合対象のCSVファイルパスのリスト
    """
    if not file_paths:
        print("結合対象のファイルが選択されていません。")
        return
    
    # 一時的に結合処理を実行
    first_file = Path(file_paths[0])
    temp_file = first_file.parent / "temp_merged.csv"
    
    try:
        # 結合処理（csv.reader/writerを使用してクオートを適切に処理）
        with open(temp_file, 'w', encoding='cp932', newline='') as outfile:
            csv_writer = csv.writer(outfile, quoting=csv.QUOTE_MINIMAL, lineterminator='\n')
            
            for i, file_path in enumerate(file_paths):
                print(f"処理中: {os.path.basename(file_path)}")
                
                try:
                    with open(file_path, 'r', encoding='cp932') as infile:
                        csv_reader = csv.reader(infile)
                        
                        for row_num, row in enumerate(csv_reader):
                            # 最初のファイルは全行書き込み、2つ目以降はヘッダー（1行目）をスキップ
                            if i == 0 or row_num > 0:
                                # 空の行をスキップ（完全に空の場合のみ）
                                if row and any(cell.strip() for cell in row):
                                    csv_writer.writerow(row)
                
                except UnicodeDecodeError:
                    # cp932で読めない場合はUTF-8で試行
                    print(f"  → UTF-8エンコーディングで再試行: {os.path.basename(file_path)}")
                    with open(file_path, 'r', encoding='utf-8') as infile:
                        csv_reader = csv.reader(infile)
                        
                        for row_num, row in enumerate(csv_reader):
                            if i == 0 or row_num > 0:
                                if row and any(cell.strip() for cell in row):
                                    csv_writer.writerow(row)
                
                except Exception as e:
                    print(f"  ⚠️ ファイル読み込みエラー ({os.path.basename(file_path)}): {e}")
                    continue
        
        print("\n結合処理完了！")
        print("出力ファイル名を選択してください...")
        
        # ファイル名選択ダイアログを表示
        selected_name = show_filename_dialog()
        
        if selected_name:
            # 出力フォルダを取得
            output_folder = get_output_folder(selected_name)
            
            if output_folder is None:
                # エラーメッセージ
                root = tk.Tk()
                root.withdraw()
                messagebox.showerror(
                    "エラー", 
                    f"出力フォルダが見つかりません。\n\n"
                    f"選択: {selected_name}\n\n"
                    f"フォルダが存在するか確認してください。"
                )
                root.destroy()
                temp_file.unlink()
                return
            
            # 最終的な出力ファイル名
            output_file = output_folder / f"{selected_name}.csv"
            
            # 既存ファイルが存在する場合は削除してから移動
            if output_file.exists():
                output_file.unlink()
                print(f"既存ファイルを削除しました: {output_file}")
            
            # 一時ファイルを最終ファイル名に移動
            shutil.move(str(temp_file), str(output_file))
            
            # 選択されたファイル名に応じて追加処理を実行
            employee_update_success = True
            db_creation_success = True
            order_data_db_creation_success = True
            order_outstanding_success = True
            outstanding_db_success = True
            purchase_order_data_db_success = True
            arrival_data_db_success = True
            shipping_db_creation_success = True
            purchase_price_db_creation_success = True

            if selected_name == "【標準】_受注":
                # 社員情報追加
                employee_update_success = update_csv_with_employee_data(output_file)

                # 受注案件状況確認DB作成
                db_creation_success = create_order_status_database(output_file)

                # 受注データDB作成（CSVと同内容）
                order_data_db_creation_success = create_order_data_database(output_file)
                
            elif selected_name == "【標準】_発注":
                # 発注残ファイル作成と発注情報DB作成
                csv_success, db_success = process_order_outstanding_csv(output_file)
                order_outstanding_success = csv_success
                outstanding_db_success = db_success

                # 発注データDB作成（CSVと同内容）
                purchase_order_data_db_success = create_purchase_order_data_database(output_file)

            elif selected_name == "【標準】_入荷":
                # 入荷データDB作成（CSVと同内容）
                arrival_data_db_success = create_arrival_data_database(output_file)

            elif selected_name == "【標準】_出荷":
                # 出荷情報DB作成
                shipping_db_creation_success = create_shipping_database(output_file)
            
            elif selected_name == "【標準】_仕入":
                # 仕入単価DB作成
                purchase_price_db_creation_success = create_purchase_price_database(output_file)
            
            # 元のファイルを削除
            deleted_count, skipped_count, failed_count = delete_source_files(file_paths)
            
            # 結果を検証
            final_size = output_file.stat().st_size
            
            print(f"\n結合完了！")
            print(f"出力ファイル: {output_file}")
            print(f"結合したファイル数: {len(file_paths)}")
            print(f"出力ファイルサイズ: {final_size:,} bytes ({final_size/1024/1024:.2f} MB)")
            
            # 行数をカウント
            try:
                with open(output_file, 'r', encoding='cp932') as f:
                    line_count = sum(1 for _ in f)
                print(f"総行数: {line_count:,} 行")
            except:
                pass
            
            # 成功メッセージ
            root = tk.Tk()
            root.withdraw()
            
            success_message = (
                f"結合が完了しました！\n\n"
                f"出力ファイル: {output_file.name}\n"
                f"ファイルサイズ: {final_size/1024/1024:.2f} MB\n"
                f"結合ファイル数: {len(file_paths)} 個\n"
            )
            
            if selected_name == "【標準】_受注":
                if employee_update_success:
                    success_message += "\n✅ 社員情報の追加も完了しました"
                else:
                    success_message += "\n⚠️ 社員情報の追加に失敗しました"
                
                if db_creation_success:
                    success_message += "\n✅ 受注案件状況確認DBの作成も完了しました"
                else:
                    success_message += "\n⚠️ 受注案件状況確認DBの作成に失敗しました"

                if order_data_db_creation_success:
                    success_message += "\n✅ 受注データDB(order_data.db)の作成も完了しました"
                else:
                    success_message += "\n⚠️ 受注データDBの作成に失敗しました"

            elif selected_name == "【標準】_発注":
                if order_outstanding_success:
                    success_message += "\n✅ 発注残ファイル(Outstanding orders.csv)の作成も完了しました"
                    success_message += "\n✅ 発注情報データベースの作成も完了しました"
                    if outstanding_db_success:
                        success_message += "\n✅ Outstanding_orders.dbの作成も完了しました"
                    else:
                        success_message += "\n⚠️ Outstanding_orders.dbの作成に失敗しました"
                else:
                    success_message += "\n⚠️ 発注残ファイル作成に失敗しました"

                if purchase_order_data_db_success:
                    success_message += "\n✅ 発注データDB(purchase_order_data.db)の作成も完了しました"
                else:
                    success_message += "\n⚠️ 発注データDBの作成に失敗しました"

            elif selected_name == "【標準】_入荷":
                if arrival_data_db_success:
                    success_message += "\n✅ 入荷データDB(arrival_data.db)の作成も完了しました"
                else:
                    success_message += "\n⚠️ 入荷データDBの作成に失敗しました"

            elif selected_name == "【標準】_出荷":
                if shipping_db_creation_success:
                    success_message += "\n✅ 出荷情報データベース(shipped.db)の作成も完了しました"
                else:
                    success_message += "\n⚠️ 出荷情報データベース作成に失敗しました"
            
            elif selected_name == "【標準】_仕入":
                if purchase_price_db_creation_success:
                    success_message += "\n✅ 仕入単価データベース(Purchase_price.db)の作成も完了しました"
                else:
                    success_message += "\n⚠️ 仕入単価データベース作成に失敗しました"
            
            # 削除処理の結果を追加
            success_message += f"\n\n元ファイル削除: {deleted_count} 件"
            if skipped_count > 0:
                success_message += f"\n削除スキップ: {skipped_count} 件"
            if failed_count > 0:
                success_message += f"\n削除失敗: {failed_count} 件"
            
            success_message += f"\n\n保存場所:\n{output_file.parent}"
            
            messagebox.showinfo("完了", success_message)
            root.destroy()
            
        else:
            # キャンセルされた場合は一時ファイルを削除
            temp_file.unlink()
            print("操作がキャンセルされました。")
        
    except Exception as e:
        # エラー時は一時ファイルを削除
        if temp_file.exists():
            temp_file.unlink()
        print(f"エラーが発生しました: {e}")
        
        # エラーメッセージ
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("エラー", f"結合処理中にエラーが発生しました:\n\n{e}")
        root.destroy()

def main():
    # コマンドライン引数からファイルパスを取得
    file_paths = sys.argv[1:]
    
    if not file_paths:
        # GUI環境でエラー表示
        root = tk.Tk()
        root.withdraw()
        messagebox.showwarning(
            "使用方法",
            "使用方法:\nファイルを選択して右クリック→送る→結合プログラム\n\n"
            "または、CSVファイルをこのプログラムにドラッグ&ドロップしてください。"
        )
        root.destroy()
        return
    
    # CSVファイルのみをフィルタリング
    csv_files = [f for f in file_paths if f.lower().endswith('.csv')]
    
    if not csv_files:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("エラー", "CSVファイルが選択されていません。")
        root.destroy()
        return
    
    print("=== CSVファイル結合プログラム（改良版） ===")
    print(f"現在のユーザー: {getpass.getuser()}")
    print(f"選択されたCSVファイル: {len(csv_files)}個")
    
    for i, file_path in enumerate(csv_files, 1):
        file_size = os.path.getsize(file_path)
        print(f"{i}. {os.path.basename(file_path)} ({file_size/1024/1024:.2f} MB)")
    
    print("\n結合を開始します...")
    print("📝 処理内容:")
    print("  ・不要なクオート(\")を除去")
    print("  ・改行コードをLF(\\n)に統一")
    print("  ・重複ヘッダーを除去")
    print("  ・空行をスキップ")
    print("  ・指定フォルダに自動保存（上書き）")
    print("  ・【標準】_受注の場合：社員情報を自動追加 + 受注案件状況確認DBを自動作成 + 受注データDB(order_data.db)を自動作成")
    print("  ・【標準】_発注の場合：発注残ファイル(Outstanding orders.csv)を自動作成 + 発注情報DBを自動作成 + 発注データDB(purchase_order_data.db)を自動作成")
    print("  ・【標準】_入荷の場合：入荷データDB(arrival_data.db)を自動作成")
    print("  ・【標準】_出荷の場合：出荷情報DB(shipped.db)を自動作成")
    print("  ・【標準】_仕入の場合：仕入単価DB(Purchase_price.db)を自動作成")
    print("  ・結合完了後、元ファイルを自動削除（削除対象外ファイルを除く）")
    print("")
    
    merge_csv_files(csv_files)

if __name__ == "__main__":
    main()