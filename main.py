import os
import sys
import json
import time
import tkinter as tk
from datetime import datetime, date
from tkinter import simpledialog, messagebox
from ui import DataCheckerImporter

def check_password(input_pw, current_date):
    return input_pw == f"zg{current_date.year}#"

def main():
    current_date = datetime.now().date()
    FREE_UNTIL = date(2026, 12, 31)

    root = tk.Tk()
    root.withdraw()

    if current_date <= FREE_UNTIL:
        root.deiconify()
        app = DataCheckerImporter(root)
        root.mainloop()
        return

    license_file = os.path.join(os.path.dirname(sys.argv[0]), "license.dat")
    
    data = {}
    if os.path.exists(license_file):
        try:
            with open(license_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
        except:
            pass
    
    data.setdefault('auth_date', None)
    data.setdefault('last_run', None)
    data.setdefault('total_errors', 0)
    data.setdefault('lock_until', None)
    
    lock_until = data.get('lock_until')
    if lock_until:
        try:
            lock_ts = float(lock_until)
            if time.time() < lock_ts:
                remain = lock_ts - time.time()
                if remain < 3600:
                    msg = f"因多次密码错误，程序已锁定，请等待 {int(remain//60)} 分 {int(remain%60)} 秒后重试。"
                else:
                    hours = int(remain // 3600)
                    minutes = int((remain % 3600) // 60)
                    msg = f"因多次密码错误，程序已锁定，请等待 {hours} 小时 {minutes} 分后重试。"
                messagebox.showerror("锁定", msg, parent=root)
                root.destroy()
                sys.exit(0)
        except:
            pass
    
    valid = False
    if data.get('auth_date') and data.get('last_run'):
        try:
            auth_date = datetime.strptime(data['auth_date'], "%Y-%m-%d").date()
            last_run = datetime.strptime(data['last_run'], "%Y-%m-%d").date()
            if current_date < last_run:
                raise ValueError("TIMEYIBEIXIUGAI")
            if (current_date - auth_date).days <= 365:
                valid = True
        except:
            pass
    
    if not valid:
        password = simpledialog.askstring("验证", f"当前日期：{current_date}\n请输入密码:", parent=root, show='*')
        if password is None:
            root.destroy()
            sys.exit(0)
        
        if not check_password(password, current_date):
            data['total_errors'] = data.get('total_errors', 0) + 1
            errors = data['total_errors']
            lock_until = None
            if errors >= 10:
                lock_until = time.time() + 86400
                msg = "密码错误已达10次，程序已锁定24小时。"
            elif errors >= 3:
                lock_until = time.time() + 180
                msg = "密码错误已达3次，程序已锁定3分钟。"
            else:
                msg = f"密码错误，还剩 {3 - errors} 次机会触发3分钟锁，{10 - errors} 次机会触发24小时锁。"
            
            if lock_until:
                data['lock_until'] = str(lock_until)
            with open(license_file, 'w', encoding='utf-8') as f:
                json.dump({
                    'auth_date': data.get('auth_date'),
                    'last_run': data.get('last_run'),
                    'total_errors': data['total_errors'],
                    'lock_until': data.get('lock_until')
                }, f)
            messagebox.showerror("错误", f"{msg}\n程序将退出。", parent=root)
            root.destroy()
            sys.exit(1)
        else:
            data['total_errors'] = 0
            data['lock_until'] = None
            if not data.get('auth_date'):
                data['auth_date'] = current_date.strftime("%Y-%m-%d")
            data['last_run'] = current_date.strftime("%Y-%m-%d")
            with open(license_file, 'w', encoding='utf-8') as f:
                json.dump(data, f)
    else:
        data['last_run'] = current_date.strftime("%Y-%m-%d")
        data['total_errors'] = 0
        data['lock_until'] = None
        with open(license_file, 'w', encoding='utf-8') as f:
            json.dump(data, f)
    
    root.deiconify()
    app = DataCheckerImporter(root)
    root.mainloop()

if __name__ == "__main__":
    main()