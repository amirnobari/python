from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from dbfread import DBF
import pandas as pd


def select_save_path(entry_widget):
    save_path = filedialog.asksaveasfilename(
        title="انتخاب مسیر ذخیره", defaultextension=".xlsx")
    entry_widget.delete(0, END)
    entry_widget.insert(END, save_path)


def select_excel_file(entry_widget):
    excel_file = filedialog.askopenfilename(title="انتخاب فایل اکسل")
    entry_widget.delete(0, END)
    entry_widget.insert(END, excel_file)


def convert_dbf_to_excel():
    # ساخت پنجره ویژوال جدید
    window = Tk()
    window.title("انتخاب فایل DBF و تنظیمات")
    window.geometry("400x300")
    window.configure(bg='#f2f2f2')  # تنظیم رنگ پس زمینه

    # دریافت مسیر فایل DBF
    dbf_file_path = filedialog.askopenfilename(title="انتخاب فایل DBF")

    # دریافت مسیر و نام فایل اکسل جدید
    excel_file_path = filedialog.asksaveasfilename(
        title="ذخیره فایل اکسل", defaultextension=".xlsx")

    if dbf_file_path and excel_file_path:
        # خواندن فایل DBF
        dbf_table = DBF(dbf_file_path)
        dataframe = pd.DataFrame(iter(dbf_table))

        # حذف ردیف‌های تکراری
        dataframe.drop_duplicates(inplace=True)

        # تبدیل ردیف‌های مورد نظر
        for index, row in dataframe.iterrows():
            date = str(row['DATE'])  # فرضا نام ستون مربوط به تاریخ DATE است
            if date.startswith('0'):
                date = '140' + date[1:]
            elif date.startswith(('7', '8', '9')):
                date = '13' + date

            # به‌روزرسانی مقدار در ردیف
            dataframe.at[index, 'DATE'] = date

        # ذخیره فایل اکسل
        dataframe.to_excel(excel_file_path, index=False)
        messagebox.showinfo("عملیات موفق", "عملیات با موفقیت انجام شد.")
        # مسیر فایل اکسل تبدیل شده
        label_file_path = Label(window, text="مسیر فایل اکسل:")
        label_file_path.pack(pady=(10, 0))
        label_file_path.configure(bg='#f2f2f2', fg='#333333', font=(
            'Arial', 12, 'bold'))  # تنظیمات رنگ و قلم
        entry_file_path = Entry(window, bg='#ffffff', fg='#333333', font=(
            'Arial', 11))  # تنظیمات رنگ و قلم
        entry_file_path.pack(pady=(0, 10))
        entry_file_path.insert(END, excel_file_path)

        # دکمه انتخاب فایل اکسل
        button_select_excel = Button(
            window, text="انتخاب فایل", command=lambda: select_excel_file(entry_file_path))
        button_select_excel.pack()

        # مسیر ذخیره
        label_save_path = Label(window, text="مسیر ذخیره:")
        label_save_path.pack(pady=(10, 0))
        label_save_path.configure(bg='#f2f2f2', fg='#333333', font=(
            'Arial', 12, 'bold'))  # تنظیمات رنگ و قلم
        entry_save_path = Entry(window, bg='#ffffff', fg='#333333', font=(
            'Arial', 11))  # تنظیمات رنگ و قلم
        entry_save_path.pack(pady=(0, 10))

        # دکمه انتخاب مسیر ذخیره
        button_select_save = Button(
            window, text="انتخاب مسیر", command=lambda: select_save_path(entry_save_path))
        button_select_save.pack()

        # دکمه اعمال
        def apply_changes():
            file_path = entry_file_path.get()
            save_path = entry_save_path.get()

            if file_path and save_path:
                df = pd.read_excel(file_path)

                # اعمال فیلترها
                # شماره کارت
                card_number = entry_card_number.get()
                filtered_df = df[df['CARD'] == int(card_number)]

                # تاریخ شروع و پایان
                start_date = entry_start_date.get()
                end_date = entry_end_date.get()
                filtered_df = filtered_df[(filtered_df['DATE'] >= int(
                    start_date)) & (filtered_df['DATE'] <= int(end_date))]

                # ذخیره فایل
                filtered_df.to_excel(save_path, index=False)
                # نمایش پیغام با استفاده از پنجره اطلاع‌رسانی
                messagebox.showinfo(
                    "عملیات موفق", "عملیات با موفقیت انجام شد.")
            else:
                messagebox.showinfo(
                    "لطفاً مسیر فایل DBF و مسیر و نام فایل اکسل جدید را مشخص کنید.")

        # تاریخ شروع
        label_start_date = Label(window, text="تاریخ شروع:")
        label_start_date.pack(pady=(10, 0))
        label_start_date.configure(bg='#f2f2f2', fg='#333333', font=(
            'Arial', 12, 'bold'))  # تنظیمات رنگ و قلم
        entry_start_date = Entry(window, bg='#ffffff', fg='#333333', font=(
            'Arial', 11))  # تنظیمات رنگ و قلم
        entry_start_date.pack(pady=(0, 10))

        # تاریخ پایان
        label_end_date = Label(window, text="تاریخ پایان:")
        label_end_date.pack(pady=(10, 0))
        label_end_date.configure(bg='#f2f2f2', fg='#333333', font=(
            'Arial', 12, 'bold'))  # تنظیمات رنگ و قلم
        entry_end_date = Entry(window, bg='#ffffff', fg='#333333', font=(
            'Arial', 11))  # تنظیمات رنگ و قلم
        entry_end_date.pack(pady=(0, 10))

        # شماره کارت
        label_card_number = Label(window, text="شماره کارت:")
        label_card_number.pack(pady=(10, 0))
        label_card_number.configure(bg='#f2f2f2', fg='#333333', font=(
            'Arial', 12, 'bold'))  # تنظیمات رنگ و قلم
        entry_card_number = Entry(window, bg='#ffffff', fg='#333333', font=(
            'Arial', 11))  # تنظیمات رنگ و قلم
        entry_card_number.pack(pady=(0, 10))

        button_apply = Button(window, text="اعمال", command=apply_changes)
        button_apply.configure(bg='#ffffff', fg='#333333', font=(
            'Arial', 12, 'bold'))  # تنظیمات رنگ و قلم
        button_apply.pack(pady=(10, 0))

        window.mainloop()
    else:
        messagebox.showinfo(
            "لطفاً مسیر فایل DBF و مسیر و نام فایل اکسل جدید را مشخص کنید.")


def convert_excel_to_excel():
    # ساخت پنجره ویژوال جدید
    window = Tk()
    window.title("انتخاب فایل اکسل و تنظیمات")
    window.geometry("400x600")
    window.configure(bg='#f2f2f2')  # تنظیم رنگ پس زمینه

    # دریافت مسیر فایل اکسل اصلی
    original_file_path = filedialog.askopenfilename(
        title="انتخاب فایل اکسل اصلی")

    if original_file_path:
        # مسیر فایل اکسل اصلی
        label_original_file_path = Label(window, text="مسیر فایل اکسل اصلی:")
        label_original_file_path.pack(pady=(10, 0))
        label_original_file_path.configure(bg='#f2f2f2', fg='#333333', font=(
            'Arial', 12, 'bold'))  # تنظیمات رنگ و قلم
        entry_original_file_path = Entry(
            window, bg='#ffffff', fg='#333333', font=('Arial', 11))  # تنظیمات رنگ و قلم
        entry_original_file_path.pack(pady=(0, 10))
        entry_original_file_path.insert(END, original_file_path)

        # دکمه انتخاب فایل اکسل اصلی
        button_select_original = Button(
            window, text="انتخاب فایل", command=lambda: select_excel_file(entry_original_file_path))
        button_select_original.pack()

        # مسیر ذخیره
        label_save_path = Label(window, text="مسیر ذخیره :")
        label_save_path.pack(pady=(10, 0))
        label_save_path.configure(bg='#f2f2f2', fg='#333333', font=(
            'Arial', 12, 'bold'))  # تنظیمات رنگ و قلم
        entry_save_path = Entry(window, bg='#ffffff', fg='#333333', font=(
            'Arial', 11))  # تنظیمات رنگ و قلم
        entry_save_path.pack(pady=(0, 10))

        # دکمه انتخاب مسیر ذخیره
        button_select_save = Button(
            window, text="انتخاب مسیر", command=lambda: select_save_path(entry_save_path))
        button_select_save.pack()

        # دکمه اعمال
        def apply_changes():
            file_path = entry_original_file_path.get()
            save_path = entry_save_path.get()

            if file_path and save_path:
                df = pd.read_excel(file_path)

                # اعمال فیلترها
                # شماره کارت
                card_number = entry_card_number.get()
                filtered_df = df[df['CARD'] == int(card_number)]

                # تاریخ شروع و پایان
                start_date = entry_start_date.get()
                end_date = entry_end_date.get()
                filtered_df = filtered_df[(filtered_df['DATE'] >= int(
                    start_date)) & (filtered_df['DATE'] <= int(end_date))]

                # ذخیره فایل
                filtered_df.to_excel(save_path, index=False)

                messagebox.showinfo(
                    "عملیات موفق", "عملیات با موفقیت انجام شد.")
            else:
                messagebox.showinfo(
                    "لطفاً مسیر فایل اکسل اصلی و مسیر ذخیره را مشخص کنید.")
        # تاریخ شروع
        label_start_date = Label(window, text="تاریخ شروع:")
        label_start_date.pack(pady=(10, 0))
        label_start_date.configure(bg='#f2f2f2', fg='#333333', font=(
            'Arial', 12, 'bold'))  # تنظیمات رنگ و قلم
        entry_start_date = Entry(window, bg='#ffffff', fg='#333333', font=(
            'Arial', 11))  # تنظیمات رنگ و قلم
        entry_start_date.pack(pady=(0, 10))

        # تاریخ پایان
        label_end_date = Label(window, text="تاریخ پایان:")
        label_end_date.pack(pady=(10, 0))
        label_end_date.configure(bg='#f2f2f2', fg='#333333', font=(
            'Arial', 12, 'bold'))  # تنظیمات رنگ و قلم
        entry_end_date = Entry(window, bg='#ffffff', fg='#333333', font=(
            'Arial', 11))  # تنظیمات رنگ و قلم
        entry_end_date.pack(pady=(0, 10))

        # شماره کارت
        label_card_number = Label(window, text="شماره کارت:")
        label_card_number.pack(pady=(10, 0))
        label_card_number.configure(bg='#f2f2f2', fg='#333333', font=(
            'Arial', 12, 'bold'))  # تنظیمات رنگ و قلم
        entry_card_number = Entry(window, bg='#ffffff', fg='#333333', font=(
            'Arial', 11))  # تنظیمات رنگ و قلم
        entry_card_number.pack(pady=(0, 10))

        button_apply = Button(window, text="اعمال", command=apply_changes)
        button_apply.configure(bg='#ffffff', fg='#333333', font=(
            'Arial', 12, 'bold'))  # تنظیمات رنگ و قلم
        button_apply.pack(pady=(10, 0))

        window.mainloop()
    else:

        messagebox.showinfo("لطفاً مسیر فایل اکسل اصلی را مشخص کنید.")


# ساخت پنجره و رابط کاربری
root = Tk()
root.title("به اکسل DBF تبدیل")
root.geometry("400x200")
root.configure(bg='#f2f2f2')  # تنظیم رنگ پس زمینه

# دکمه تبدیل DBF به اکسل
convert_dbf_button = Button(
    root, text="به اکسل DBF تبدیل", command=convert_dbf_to_excel)
convert_dbf_button.configure(bg='#ffffff', fg='#333333', font=(
    'Arial', 12, 'bold'))  # تنظیمات رنگ و قلم
convert_dbf_button.pack(pady=(10, 0))

# دکمه تبدیل اکسل به اکسل
convert_excel_button = Button(
    root, text="تبدیل اکسل به اکسل", command=convert_excel_to_excel)
convert_excel_button.configure(bg='#ffffff', fg='#333333', font=(
    'Arial', 12, 'bold'))  # تنظیمات رنگ و قلم
convert_excel_button.pack(pady=(10, 0))

root.mainloop()
