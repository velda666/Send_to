import sys
import os
import pyperclip
import time
from datetime import datetime
import tkinter as tk
from tkinter import messagebox

# 旧担当者案件も残っている可能性があるため、片平両名・荒井氏・酒井氏のコードは暫く残しておく
addition_dict = {
   "1001": "_Oｶﾀﾋﾗｾ",
   "1002": "_Tｶﾀﾋﾗﾌﾞ",
   "1003": "_Oﾐﾀﾆ",
   "1004": "_Oﾀﾅｶ",
   "1005": "_Wﾖｺﾔﾏ",
   "1006": "_Wﾌﾙﾐﾔ",
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
            
            if special_text_key == "1003" and check_value == "BJF3C":
                special_text = "1003_Oﾐﾀﾆ_COSCO"
            elif special_text_key == "1006" and (
                check_value in ["63G50", "63G51", "BJF3C"] or 
                seventh_eighth_value in ["A104-11", "A104-12", "A114-11", "A114-CS", "A114-DL", 
                                       "A114-GZ", "A114-HK", "A114-QD", "A114-SG", "A114-SH", 
                                       "A114-SZ", "A114-TJ", "A114-XM"]
            ):
                special_text = "1006_Wﾌﾙﾐﾔ_ｶｲｶﾞｲ"
            elif len(underscore_positions) >= 8 and "A042" in file_name[underscore_positions[7] + 1:underscore_positions[8]]:
                special_text = "1008_ﾌｼﾞｳ_SW"
            elif len(underscore_positions) >= 8 and seventh_eighth_value == "A056-11":
                special_text = "1008_ﾌｼﾞｳ_ﾎﾟｰﾄ_ﾋﾗﾏﾂ"
            elif len(underscore_positions) >= 8 and "A195" in file_name[underscore_positions[7] + 1:underscore_positions[8]]:
                special_text = "1025_Tｷﾀﾉ_ｷｮｳﾄﾞｳ"
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
   if len(sys.argv) > 1:
       file_path = sys.argv[1]
       copy_partial_filename_and_path(file_path)
       app = FileMoveTool(file_path)
       app.run()
   else:
       print("No file path provided.")