import sys
import os
import getpass
import shutil
import subprocess
import winreg
import pyperclip
import time
from datetime import datetime
import tkinter as tk
from tkinter import messagebox

# アプリケーションバージョン
APP_VERSION = "1.1.8"
APP_NAME = "発注ファイル_必要情報コピー"

# 旧担当者案件も残っている可能性があるため、片平両名・荒井氏・酒井氏のコードは暫く残しておく
addition_dict = {
   "1001": "_Oｶﾀﾋﾗｾ",
   "1002": "_Tｶﾀﾋﾗﾌﾞ",
   "1003": "_Oﾐﾀﾆ",
   "1004": "_Oﾀﾅｶ",
   "1005": "_Wﾖｺﾔﾏ",
   "1006": "_Oﾌﾙﾐﾔ",
   "1007": "_Kﾐﾅﾐﾊﾗ",
   "1008": "_Tﾌｼﾞｳ",
   "1009": "_Tｱﾗｲ",
   "1010": "_Wﾌﾙﾐﾔ",
   "1014": "_Wｳｴﾉ",
   "1016": "_Oｵｶﾑﾗ",
   "1025": "_Tｷﾀﾉ",
   "1028": "_Oｲﾄﾞ",
   "1030": "_Oｶﾅﾓﾘ"
   
}

def copy_to_clipboard(text):
   pyperclip.copy(text)
   print(f"Copied to clipboard: {text}")
   time.sleep(0.5)

def copy_partial_filename_and_path(file_path):
    file_name = os.path.splitext(os.path.basename(file_path))[0]
    underscore_positions = [i for i, char in enumerate(file_name) if char == '_']
    
    if len(underscore_positions) >= 3:
        partial_name = file_name[underscore_positions[1] + 1:underscore_positions[2]]
    else:
        partial_name = ""

    if "nosales" in file_name:
        special_text = "在庫"
    else:
        if len(underscore_positions) >= 6:
            special_text_key = file_name[underscore_positions[4] + 1:underscore_positions[5]]
            check_value = file_name[underscore_positions[2] + 1:underscore_positions[3]]
            
            # 7番目と8番目のアンダースコア間の値を取得（A114-XMの位置）
            seventh_eighth_value = ""
            if len(underscore_positions) >= 8:
                seventh_eighth_value = file_name[underscore_positions[7] + 1:underscore_positions[8]]
            # 倉庫の要望でCOSCOはミタニ表記。BJF3Cは担当者コード問わず仕向地により分岐
            if check_value == "BJF3C":
                if "A220" in seventh_eighth_value:
                    special_text = "1003_Oﾐﾀﾆ_COSCO"
                elif "A114" in seventh_eighth_value:
                    special_text = "1010_Wﾌﾙﾐﾔ_ｶｲｶﾞｲ_YT"
                elif "A113" in seventh_eighth_value or "A103" in seventh_eighth_value:
                    special_text = "1014_Wｳｴﾉ_COSCO"
                else:
                    special_text = "1003_Oﾐﾀﾆ_COSCO"
            # A104/A114は担当者コード問わず古宮海外
            elif "A104" in seventh_eighth_value or "A114" in seventh_eighth_value:
                special_text = "1010_Wﾌﾙﾐﾔ_ｶｲｶﾞｲ"
            # 国内古宮コード（63G50/63G51の場合）
            elif special_text_key == "1006" and check_value in ["63G50", "63G51"]:
                special_text = "1006_Wﾌﾙﾐﾔ_ｶｲｶﾞｲ"
            elif special_text_key == "1010" and check_value in ["63G50", "63G51"]:
                special_text = "1010_Wﾌﾙﾐﾔ_ｶｲｶﾞｲ"
            elif seventh_eighth_value in ["A112-11", "A111-11", "A111-12", "A111-13", "A111-14"]:
                special_text = "1014_Wｳｴﾉ_ｶｲｶﾞｲ"
            elif len(underscore_positions) >= 8 and "A042" in file_name[underscore_positions[7] + 1:underscore_positions[8]]:
                special_text = "1008_ﾌｼﾞｳ_SW"
            elif len(underscore_positions) >= 8 and seventh_eighth_value == "A056-11":
                special_text = "1008_ﾌｼﾞｳ_ﾎﾟｰﾄ_ﾋﾗﾏﾂ"
            elif len(underscore_positions) >= 8 and "A195" in file_name[underscore_positions[7] + 1:underscore_positions[8]]:
                special_text = "1025_Tｷﾀﾉ_ｷｮｳﾄﾞｳ"
            elif "A220" in seventh_eighth_value and check_value == "63047":
                special_text = "1003_Oﾐﾀﾆ"
            else:
                special_text = special_text_key + addition_dict.get(special_text_key, "")
        else:
            special_text = ""

    full_path = file_path

    if len(underscore_positions) >= 8:
        additional_text1 = file_name[underscore_positions[5] + 1:underscore_positions[6]]
        additional_text2 = file_name[underscore_positions[6] + 1:underscore_positions[7]]
        
        copy_to_clipboard(partial_name)
        copy_to_clipboard(full_path)
        copy_to_clipboard(special_text)
        copy_to_clipboard(additional_text1)
        copy_to_clipboard(additional_text2)
    else:
        copy_to_clipboard(partial_name)
        copy_to_clipboard(full_path)
        copy_to_clipboard(special_text)

def move_file(file_path):
   directory = os.path.dirname(file_path)
   file_name = os.path.basename(file_path)
   timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
   new_file_name = f"★DONE_{timestamp}_" + file_name
   target_directory = os.path.join(directory, "発注済")

   if not os.path.exists(target_directory):
       os.makedirs(target_directory)

   target_path = os.path.join(target_directory, new_file_name)
   os.rename(file_path, target_path)
   print(f"Moved file to: {target_path}")

# -------------------------------------------------------
# アップデート関連関数
# -------------------------------------------------------
REGISTRY_KEY_PATH = r"SOFTWARE\TohoYanmar\OrderFileCopyApp"

def get_update_folder_path():
    """アップデート用フォルダのパスを取得する"""
    username = getpass.getuser()
    candidate_paths = [
        f"C:\\Users\\{username}\\OneDrive - 東邦ヤンマーテック株式会社\\CR推進本部フォルダ\\06_社内管理資料\\miraimiru移行関連\\フォルダ共有テスト\\業務用pythonアプリ最新版\\発注ファイル_必要情報コピー\\update",
        f"C:\\Users\\{username}\\東邦ヤンマーテック株式会社\\CR推進本部フォルダ\\06_社内管理資料\\miraimiru移行関連\\フォルダ共有テスト\\業務用pythonアプリ最新版\\発注ファイル_必要情報コピー\\update",
        f"C:\\Users\\{username}\\東邦ヤンマーテック株式会社\\CR推進本部 - CR推進本部フォルダ\\06_社内管理資料\\miraimiru移行関連\\フォルダ共有テスト\\業務用pythonアプリ最新版\\発注ファイル_必要情報コピー\\update",
    ]
    for path in candidate_paths:
        if os.path.exists(path):
            return path
    return None

def get_registry_value(value_name, default=None):
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, REGISTRY_KEY_PATH, 0, winreg.KEY_READ)
        value, _ = winreg.QueryValueEx(key, value_name)
        winreg.CloseKey(key)
        return value
    except FileNotFoundError:
        return default
    except Exception:
        return default

def set_registry_value(value_name, value):
    try:
        key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, REGISTRY_KEY_PATH)
        winreg.SetValueEx(key, value_name, 0, winreg.REG_SZ, str(value))
        winreg.CloseKey(key)
    except Exception:
        pass

def get_skip_count(version):
    skip_info = get_registry_value("SkipInfo", "")
    if skip_info:
        try:
            parts = skip_info.split("|")
            if len(parts) == 2 and parts[0] == version:
                return int(parts[1])
        except Exception:
            pass
    return 0

def set_skip_count(version, count):
    set_registry_value("SkipInfo", f"{version}|{count}")

def reset_skip_count():
    set_registry_value("SkipInfo", "")

def compare_versions(version1, version2):
    """version1 > version2 なら正、等しければ0、小さければ負を返す"""
    def parse_version(v):
        return [int(x) for x in v.split('.')]
    v1_parts = parse_version(version1)
    v2_parts = parse_version(version2)
    for v1_val, v2_val in zip(v1_parts, v2_parts):
        if v1_val > v2_val:
            return 1
        elif v1_val < v2_val:
            return -1
    return 0

def get_latest_version():
    update_folder = get_update_folder_path()
    if not update_folder:
        return None
    version_file = os.path.join(update_folder, "version.txt")
    if not os.path.exists(version_file):
        return None
    try:
        with open(version_file, 'r', encoding='utf-8') as f:
            return f.read().strip()
    except Exception:
        return None

def perform_update():
    """アップデートを実行する"""
    update_folder = get_update_folder_path()
    if not update_folder:
        messagebox.showerror("エラー", "アップデート用フォルダが見つかりません。")
        return False

    is_frozen = getattr(sys, 'frozen', False)
    if is_frozen:
        current_script = sys.executable
    else:
        current_script = os.path.abspath(__file__)

    update_source = os.path.join(update_folder, os.path.basename(current_script))
    if not os.path.exists(update_source):
        messagebox.showerror("エラー", f"アップデートファイルが見つかりません:\n{update_source}")
        return False

    try:
        if not is_frozen:
            shutil.copy2(update_source, current_script)
            reset_skip_count()
            messagebox.showinfo("アップデート完了", "アップデートが完了しました。\nアプリケーションを再起動します。")
            subprocess.Popen([sys.executable, current_script, "--just-updated"] + sys.argv[1:])
            return True
        else:
            batch_content = f'''@echo off
timeout /t 2 /nobreak > nul
copy /Y "{update_source}" "{current_script}"
start "" "{current_script}" --just-updated
del "%~f0"
'''
            batch_path = os.path.join(os.path.dirname(current_script), "update_temp.bat")
            with open(batch_path, 'w', encoding='shift_jis') as f:
                f.write(batch_content)
            reset_skip_count()
            subprocess.Popen(batch_path, shell=True)
            return True
    except Exception as e:
        messagebox.showerror("エラー", f"アップデート中にエラーが発生しました:\n{str(e)}")
        return False

def check_for_updates(root=None):
    """起動時にアップデートをチェックする"""
    if "--just-updated" in sys.argv:
        return False

    latest_version = get_latest_version()
    if latest_version is None:
        return False

    if compare_versions(latest_version, APP_VERSION) <= 0:
        return False

    skip_count = get_skip_count(latest_version)

    if skip_count >= 2:
        messagebox.showinfo(
            "アップデート必須",
            f"新しいバージョン {latest_version} が利用可能です。\n"
            f"現在のバージョン: {APP_VERSION}\n\n"
            "アップデートを延期できる回数を超えました。\n"
            "アップデートを実行します。"
        )
        if perform_update():
            if root:
                root.destroy()
            sys.exit(0)
    else:
        remaining = 2 - skip_count
        result = messagebox.askyesno(
            "アップデート確認",
            f"新しいバージョン {latest_version} が利用可能です。\n"
            f"現在のバージョン: {APP_VERSION}\n\n"
            f"今すぐアップデートしますか？\n\n"
            f"（「いいえ」を選択した場合、残り{remaining}回まで延期できます）"
        )
        if result:
            if perform_update():
                if root:
                    root.destroy()
                sys.exit(0)
        else:
            set_skip_count(latest_version, skip_count + 1)

    return False

# -------------------------------------------------------

class FileMoveTool:
   def __init__(self, file_path):
       self.file_path = file_path
       self.root = tk.Tk()
       self.root.title("ファイル移動ツール")
       self.root.geometry("400x200")
       self.root.resizable(False, False)

       # ウィンドウを中央に配置
       screen_width = self.root.winfo_screenwidth()
       screen_height = self.root.winfo_screenheight()
       x = (screen_width - 400) // 2
       y = (screen_height - 200) // 2
       self.root.geometry(f"400x200+{x}+{y}")

       # メッセージラベル
       label = tk.Label(
           self.root,
           text="PO3発注後にボタンをクリックしてファイルを移動して下さい。",
           font=("", 10)
       )
       label.pack(pady=20)

       # 移動ボタン
       button = tk.Button(
           self.root,
           text="発注済みファイルを移動",
           command=self.move_button_click,
           bg="blue",
           fg="white",
           width=20,
           height=2
       )
       button.pack(pady=20)

   def move_button_click(self):
       move_file(self.file_path)
       messagebox.showinfo("完了", "ファイルが発注済みフォルダに移動されました。")
       self.root.destroy()

   def run(self):
       self.root.mainloop()

if __name__ == "__main__":
   # アップデートチェック
   temp_root = tk.Tk()
   temp_root.withdraw()
   check_for_updates(temp_root)
   try:
       temp_root.destroy()
   except Exception:
       pass

   if len(sys.argv) > 1:
       # --just-updated フラグを除いた引数からファイルパスを取得
       args = [a for a in sys.argv[1:] if a != "--just-updated"]
       if args:
           file_path = args[0]
           copy_partial_filename_and_path(file_path)
           app = FileMoveTool(file_path)
           app.run()
   else:
       print("No file path provided.")