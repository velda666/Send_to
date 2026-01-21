import sys
import os
import shutil
import getpass
from pathlib import Path
import tkinter as tk
from tkinter import messagebox
import pandas as pd
import glob

class CSVFileProcessor:
    def __init__(self):
        # 現在のユーザー名を取得
        self.username = getpass.getuser()
        
        # フォルダパスの候補を定義
        self.folder_candidates = [
            f"C:\\Users\\{self.username}\\OneDrive - 東邦ヤンマーテック株式会社\\CR推進本部フォルダ\\06_社内管理資料\\miraimiru移行関連\\フォルダ共有テスト\\受渡場所一覧",
            f"C:\\Users\\{self.username}\\東邦ヤンマーテック株式会社\\CR推進本部フォルダ\\06_社内管理資料\\miraimiru移行関連\\フォルダ共有テスト\\受渡場所一覧",
            f"C:\\Users\\{self.username}\\東邦ヤンマーテック株式会社\\CR推進本部 - CR推進本部フォルダ\\06_社内管理資料\\miraimiru移行関連\\フォルダ共有テスト\\受渡場所一覧"
        ]
        
        # 存在するフォルダパスを確認
        self.target_folder = self.find_existing_folder()
    
    def find_existing_folder(self):
        """存在するフォルダパスを見つける"""
        for folder_path in self.folder_candidates:
            if os.path.exists(folder_path):
                return folder_path
        return None
    
    def show_message(self, title, message, is_error=False):
        """メッセージボックスを表示"""
        try:
            # Tkinterウィンドウを非表示にする
            root = tk.Tk()
            root.withdraw()
            
            if is_error:
                messagebox.showerror(title, message)
            else:
                messagebox.showinfo(title, message)
            
            root.destroy()
        except:
            # GUIが使えない場合はコンソールに出力
            print(f"{title}: {message}")
    
    def validate_file(self, file_path):
        """ファイルの検証"""
        if not os.path.exists(file_path):
            return False, "指定されたファイルが存在しません。"
        
        if not file_path.lower().endswith('.csv'):
            return False, "CSVファイルを指定してください。"
        
        if not self.target_folder:
            return False, f"対象フォルダが見つかりません。\n以下のフォルダが存在するか確認してください:\n\n" + "\n".join(self.folder_candidates)
        
        return True, "OK"
    
    def process_csv_data(self, csv_file_path):
        """CSVデータの処理（列の結合とゼロパディング）"""
        try:
            # 作業ディレクトリを対象フォルダに変更
            original_dir = os.getcwd()
            os.chdir(self.target_folder)
            
            # "受渡場所"を含むCSVファイルを検索
            csv_files = glob.glob('*受渡場所*.csv')
            if not csv_files:
                raise FileNotFoundError('"受渡場所"を含むCSVファイルが見つかりませんでした。')
            
            # 最初に見つかったファイルを使用
            target_file = csv_files[0]
            
            # CSVファイルの読み込み（cp932エンコーディング）
            df = pd.read_csv(target_file, encoding='cp932')
            
            # データが空でないかチェック
            if df.empty:
                raise ValueError("CSVファイルにデータが含まれていません。")
            
            # 必要な列が存在するかチェック
            required_columns = ['得意先コード', '得意先枝番']
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                raise ValueError(f"必要な列が見つかりません: {', '.join(missing_columns)}")
            
            # 新しい列の作成（得意先コード-得意先枝番）
            combined_value = df['得意先コード'].astype(str) + '-' + df['得意先枝番'].astype(str)
            
            # データフレームの先頭（A列）に新しい列を挿入
            df.insert(0, '得意先コード-得意先枝番', combined_value)
            
            # 6列目の値をゼロパディング（2桁の数字に変換）
            if len(df.columns) >= 6:
                column_name = df.columns[5]  # 6列目のカラム名を取得
                df[column_name] = df[column_name].astype(str).str.zfill(2)
            else:
                print("警告: 6列目が存在しないため、ゼロパディング処理をスキップしました。")
            
            # 一時的な名前で保存
            temp_file = 'temp_output.csv'
            df.to_csv(temp_file, encoding='cp932', index=False)
            
            # 既存の'受渡場所.csv'が存在する場合は削除
            if os.path.exists('受渡場所.csv'):
                os.remove('受渡場所.csv')
            
            # ファイル名を'受渡場所.csv'にリネーム
            os.rename(temp_file, '受渡場所.csv')
            
            # 元のファイルが'受渡場所.csv'と異なる場合は削除
            if target_file != '受渡場所.csv':
                os.remove(target_file)
            
            # 作業ディレクトリを元に戻す
            os.chdir(original_dir)
            
            return True, "CSVデータの処理が完了しました。"
            
        except Exception as e:
            # 作業ディレクトリを元に戻す
            os.chdir(original_dir)
            return False, f"CSVデータ処理中にエラーが発生しました:\n{str(e)}"
    
    def process_file(self, source_file_path):
        """ファイルの処理（移動とリネーム + データ処理）"""
        try:
            # ファイル検証
            is_valid, message = self.validate_file(source_file_path)
            if not is_valid:
                self.show_message("エラー", message, True)
                return False
            
            # 対象ファイルパス
            target_file_path = os.path.join(self.target_folder, "受渡場所.csv")
            
            # 既存ファイルがある場合はバックアップ作成
            backup_created = False
            if os.path.exists(target_file_path):
                backup_path = os.path.join(self.target_folder, "受渡場所_backup.csv")
                shutil.copy2(target_file_path, backup_path)
                backup_created = True
            
            # ファイルをコピーしてリネーム
            shutil.copy2(source_file_path, target_file_path)
            
            # 元ファイルを削除（移動を完了）
            os.remove(source_file_path)
            
            # CSVデータの処理を実行
            data_success, data_message = self.process_csv_data(target_file_path)
            
            # 結果メッセージ
            success_message = f"ファイルの移動が完了しました。\n\n"
            success_message += f"元ファイル: {os.path.basename(source_file_path)}\n"
            success_message += f"移動先: {target_file_path}\n"
            if backup_created:
                success_message += f"\n※既存ファイルはバックアップされました\n"
            
            if data_success:
                success_message += f"\n✓ {data_message}\n"
                success_message += "・得意先コード-得意先枝番列を先頭に追加\n"
                success_message += "・6列目の値をゼロパディング（2桁）に変換"
            else:
                success_message += f"\n⚠ データ処理でエラーが発生しました:\n{data_message}"
            
            self.show_message("完了", success_message, not data_success)
            return True
            
        except Exception as e:
            self.show_message("エラー", f"ファイル処理中にエラーが発生しました:\n{str(e)}", True)
            return False

def main():
    """メイン処理"""
    # pandasの依存関係チェック
    try:
        import pandas as pd
    except ImportError:
        processor = CSVFileProcessor()
        processor.show_message("エラー", 
                             "pandasライブラリが必要です。\n\n"
                             "以下のコマンドでインストールしてください:\n"
                             "pip install pandas", True)
        return
    
    processor = CSVFileProcessor()
    
    # コマンドライン引数をチェック
    if len(sys.argv) < 2:
        processor.show_message("使用方法", 
                             "このプログラムは以下の方法で使用してください:\n\n"
                             "1. CSVファイルをこのプログラムのアイコンにドラッグ&ドロップ\n"
                             "2. CSVファイルを右クリック > 送る > このプログラム\n\n"
                             "対象: cp932エンコードのCSVファイル\n\n"
                             "処理内容:\n"
                             "・ファイルを指定フォルダに移動してリネーム\n"
                             "・得意先コード-得意先枝番列を先頭に追加\n"
                             "・6列目の値をゼロパディング（2桁）")
        return
    
    # 引数からファイルパスを取得（複数ファイルの場合は最初のみ処理）
    file_path = sys.argv[1]
    
    # ファイル処理実行
    processor.process_file(file_path)

if __name__ == "__main__":
    main()