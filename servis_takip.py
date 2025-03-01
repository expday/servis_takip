import tkinter as tk
from tkinter import ttk, messagebox, filedialog, Menu
import sqlite3
import datetime
import os
import shutil
from tkcalendar import DateEntry
import logging
import json
import pandas as pd
import platform
import re

# Global ayarlar
SETTINGS_FILE = "settings.json"
DEFAULT_SETTINGS = {
    "log_enabled": True,
    "log_file": "servis_takip.log"
}

# Ayarları yükle
def load_settings():
    global settings
    if os.path.exists(SETTINGS_FILE):
        with open(SETTINGS_FILE, "r") as f:
            settings = json.load(f)
    else:
        settings = DEFAULT_SETTINGS.copy()
        save_settings()
    return settings

# Ayarları kaydet
def save_settings():
    with open(SETTINGS_FILE, "w") as f:
        json.dump(settings, f, indent=4)

# Loglama ayarları
def configure_logging():
    if settings["log_enabled"]:
        logging.basicConfig(filename=settings["log_file"], level=logging.INFO,
                            format="%(asctime)s - %(levelname)s - %(message)s")
    else:
        logging.basicConfig(level=logging.CRITICAL)

# İlk ayarları yükle ve loglamayı yapılandır
settings = load_settings()
configure_logging()

class DatabaseManager:
    def __init__(self, db_name="servis_takip.db"):
        self.conn = sqlite3.connect(db_name)
        self.cursor = self.conn.cursor()
        self.create_tables()

    def create_tables(self):
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS cihazlar (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                barkod_no TEXT NOT NULL,
                bolge TEXT, personel_ad_soyad TEXT, personel_sicil_no TEXT,
                cihaz_tipi TEXT, cihaz_seri_no TEXT, servis_gonderim_tarihi TEXT,
                servis_gelme_tarihi TEXT, cihaz_durumu TEXT, aciklama TEXT,
                cihaz_belgeleri TEXT
            )
        ''')
        self.conn.commit()

    def insert_cihaz(self, veriler):
        try:
            veriler["cihaz_belgeleri"] = json.dumps(veriler["cihaz_belgeleri"])
            self.cursor.execute('''
                INSERT INTO cihazlar (
                    barkod_no, bolge, personel_ad_soyad, personel_sicil_no, 
                    cihaz_tipi, cihaz_seri_no, servis_gonderim_tarihi, 
                    servis_gelme_tarihi, cihaz_durumu, aciklama, cihaz_belgeleri
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                veriler['barkod_no'], veriler['bolge'], veriler['personel_ad_soyad'], 
                veriler['personel_sicil_no'], veriler['cihaz_tipi'], veriler['cihaz_seri_no'], 
                veriler['servis_gonderim_tarihi'], veriler['servis_gelme_tarihi'], 
                veriler['cihaz_durumu'], veriler['aciklama'], veriler['cihaz_belgeleri']
            ))
            self.conn.commit()
            if settings["log_enabled"]:
                logging.info(f"Cihaz kaydedildi: {veriler['barkod_no']} (Tarih: {veriler['servis_gonderim_tarihi']})")
            return self.cursor.lastrowid
        except sqlite3.Error as e:
            if settings["log_enabled"]:
                logging.error(f"Veritabanı hatası (insert_cihaz): {e}")
            return False

    def update_cihaz(self, veriler, id):
        try:
            veriler["cihaz_belgeleri"] = json.dumps(veriler["cihaz_belgeleri"])
            self.cursor.execute('''
                UPDATE cihazlar SET 
                    barkod_no = ?, bolge = ?, personel_ad_soyad = ?, personel_sicil_no = ?, 
                    cihaz_tipi = ?, cihaz_seri_no = ?, servis_gonderim_tarihi = ?, 
                    servis_gelme_tarihi = ?, cihaz_durumu = ?, aciklama = ?, 
                    cihaz_belgeleri = ? 
                WHERE id = ?
            ''', (
                veriler['barkod_no'], veriler['bolge'], veriler['personel_ad_soyad'], 
                veriler['personel_sicil_no'], veriler['cihaz_tipi'], veriler['cihaz_seri_no'], 
                veriler['servis_gonderim_tarihi'], veriler['servis_gelme_tarihi'], 
                veriler['cihaz_durumu'], veriler['aciklama'], veriler['cihaz_belgeleri'], id
            ))
            self.conn.commit()
            if settings["log_enabled"]:
                logging.info(f"Cihaz güncellendi: ID {id}")
            return True
        except sqlite3.Error as e:
            if settings["log_enabled"]:
                logging.error(f"Güncelleme hatası (update_cihaz): {e}")
            return False

    def delete_cihaz(self, id):
        self.cursor.execute("DELETE FROM cihazlar WHERE id = ?", (id,))
        self.conn.commit()
        if settings["log_enabled"]:
            logging.info(f"Cihaz silindi: ID {id}")

    def fetch_cihaz_by_id(self, id):
        self.cursor.execute("SELECT * FROM cihazlar WHERE id = ?", (id,))
        cihaz = self.cursor.fetchone()
        if cihaz:
            cihaz = list(cihaz)
            cihaz[11] = json.loads(cihaz[11]) if cihaz[11] else []
            return tuple(cihaz)
        return None

    def fetch_cihazlar_by_barkod(self, barkod_no):
        self.cursor.execute("SELECT * FROM cihazlar WHERE barkod_no = ?", (barkod_no,))
        rows = self.cursor.fetchall()
        return [tuple(list(row)[:-1] + [json.loads(row[-1]) if row[-1] else []]) for row in rows]

    def fetch_all(self, filtre=""):
        if filtre:
            self.cursor.execute("SELECT * FROM cihazlar WHERE barkod_no LIKE ? OR personel_ad_soyad LIKE ?",
                               (f"%{filtre}%", f"%{filtre}%"))
        else:
            self.cursor.execute("SELECT * FROM cihazlar")
        rows = self.cursor.fetchall()
        return [tuple(list(row)[:-1] + [json.loads(row[-1]) if row[-1] else []]) for row in rows]

    def advanced_search(self, filtreler):
        query = "SELECT * FROM cihazlar WHERE 1=1"
        params = []
        if filtreler.get("barkod_no"):
            query += " AND barkod_no LIKE ?"
            params.append(f"%{filtreler['barkod_no']}%")
        if filtreler.get("cihaz_durumu"):
            query += " AND cihaz_durumu = ?"
            params.append(filtreler["cihaz_durumu"])
        self.cursor.execute(query, params)
        rows = self.cursor.fetchall()
        return [tuple(list(row)[:-1] + [json.loads(row[-1]) if row[-1] else []]) for row in rows]

    def close(self):
        self.conn.close()

class ServisTakipUygulamasi:
    def __init__(self):
        self.root = tk.Tk()
        os.makedirs("belgeler", exist_ok=True)
        self.db = DatabaseManager()
        self.setup_main_window()

    def create_input_fields(self, frame, alanlar):
        self.entries = {}
        for i, (label, key, widget_type, *args) in enumerate(alanlar):
            ttk.Label(frame, text=label).grid(row=i, column=0, sticky="e", padx=5, pady=5)
            if widget_type == DateEntry:
                entry = DateEntry(frame, date_pattern="dd.mm.yyyy", width=15)
                entry.set_date(datetime.datetime.now())
            elif widget_type == ttk.Combobox:
                entry = ttk.Combobox(frame, values=args[0])
                entry.set(args[0][0])
            elif widget_type == tk.Text:
                entry = widget_type(frame, height=1, width=40)
            else:
                entry = widget_type(frame)
            entry.grid(row=i, column=1, sticky="w", padx=5, pady=5)
            self.entries[key] = entry

    def setup_main_window(self):
        self.root.title("Servis Takip Sistemi")
        self.root.geometry("1100x900")
        self.root.configure(bg="#f0f0f0")

        try:
            if platform.system() == "Windows":
                self.root.iconbitmap("app_icon.ico")
            else:
                icon = tk.PhotoImage(file="app_icon.png")
                self.root.iconphoto(True, icon)
        except Exception as e:
            if settings["log_enabled"]:
                logging.warning(f"İkon yüklenirken hata oluştu: {e}")

        # Menü çubuğu
        menubar = Menu(self.root)
        self.root.config(menu=menubar)
        settings_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Ayarlar", menu=settings_menu)
        settings_menu.add_command(label="Ayarları Aç", command=self.show_settings)

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TButton", padding=6, relief="flat")
        style.map("TButton", background=[("active", "#0056b3")])
        style.configure("Vertical.TScrollbar", gripcount=0, arrowsize=0, background="#f0f0f0")

        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill="both", expand=True)

        ttk.Label(main_frame, text="Cihaz Servis Takip Sistemi", font=("Helvetica", 16, "bold")).pack(pady=10)

        search_frame = ttk.Frame(main_frame)
        search_frame.pack(fill="x", pady=5)
        self.search_var = tk.StringVar()
        ttk.Entry(search_frame, textvariable=self.search_var, width=30).pack(side="left", padx=5)
        ttk.Button(search_frame, text="Ara", command=self.simple_search).pack(side="left")

        input_frame = ttk.LabelFrame(main_frame, text="Cihaz Bilgileri", padding="10")
        input_frame.pack(fill="x", pady=5)
        alanlar = [
            ("Bölge:", "bolge", ttk.Entry),
            ("Personel Ad Soyad:", "personel_ad_soyad", ttk.Entry),
            ("Personel Sicil No:", "personel_sicil_no", ttk.Entry),
            ("Cihaz Tipi:", "cihaz_tipi", ttk.Combobox, ["Laptop", "SIM Kart", "Tablet", "El Terminali","Masaüstü Bilgisayar","PC Monitör","Mobil Yazıcı","PC Yazıcı","UPS"]),
            ("Barkod No*:", "barkod_no", ttk.Entry),
            ("Cihaz Seri No:", "cihaz_seri_no", ttk.Entry),
            ("Servis Gönderim Tarihi:", "servis_gonderim_tarihi", DateEntry),
            ("Servis Gelme Tarihi:", "servis_gelme_tarihi", DateEntry),
            ("Cihaz Durumu:", "cihaz_durumu", ttk.Combobox, ["Serviste", "Servise Gönderildi", "Tamir edildi","Tamir olmuyor","Hurda"]),
            ("Açıklama:", "aciklama", tk.Text),
        ]
        self.create_input_fields(input_frame, alanlar)

        self.entries["aciklama"].config(width=40, height=1)
        self.entries["aciklama"].config(bd=2, relief="sunken", font=("Helvetica", 10), background="#ffffff", foreground="#333333")
        self.entries["aciklama"].bind("<KeyRelease>", self.check_aciklama_length)

        file_frame = ttk.LabelFrame(main_frame, text="Cihaz Belgeleri", padding="10")
        file_frame.pack(fill="x", pady=5)
        self.dosya_label = ttk.Label(file_frame, text="Henüz dosya seçilmedi")
        self.dosya_label.pack()
        ttk.Button(file_frame, text="Dosyaları Yükle", command=self.dosyalar_sec).pack(side="left", padx=5, pady=5)
        ttk.Button(file_frame, text="Belgeleri Göster", command=self.show_belgeler).pack(side="left", padx=5, pady=5)
        self.secilen_dosyalar = []

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="Kaydet", command=self.cihaz_kaydet, style="Success.TButton").pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Sorgula", command=self.show_advanced_search, style="Info.TButton").pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Güncelle", command=self.durum_guncelle, style="Warning.TButton").pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Sil", command=self.cihaz_sil, style="Danger.TButton").pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Tüm Cihazları Göster", command=self.tum_cihazlari_listele, style="Primary.TButton").pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Excel'e Aktar", command=self.export_to_excel).pack(side="left", padx=5)

        tree_frame = ttk.Frame(main_frame)
        tree_frame.pack(fill="both", expand=True, pady=10)

        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", style="Vertical.TScrollbar")
        scrollbar.pack(side="right", fill="y")

        self.tree = ttk.Treeview(tree_frame, columns=("ID", "Barkod", "Bolge", "Personel", "Sicil", "Tip", "Seri", "Gonderim", "Gelme", "Durum", "Aciklama", "Belge"), 
                                 show="headings", yscrollcommand=scrollbar.set)
        self.tree.heading("ID", text="Kayıt ID", anchor="center")
        self.tree.heading("Barkod", text="Barkod No", anchor="center")
        self.tree.heading("Bolge", text="Bölge", anchor="center")
        self.tree.heading("Personel", text="Personel Ad Soyad", anchor="center")
        self.tree.heading("Sicil", text="Sicil No", anchor="center")
        self.tree.heading("Tip", text="Cihaz Tipi", anchor="center")
        self.tree.heading("Seri", text="Seri No", anchor="center")
        self.tree.heading("Gonderim", text="Gönderim Tarihi", anchor="center")
        self.tree.heading("Gelme", text="Gelme Tarihi", anchor="center")
        self.tree.heading("Durum", text="Durum", anchor="center")
        self.tree.heading("Aciklama", text="Açıklama", anchor="center")
        self.tree.heading("Belge", text="Belge Sayısı", anchor="center")
        self.tree.column("ID", width=60, anchor="center")
        self.tree.column("Barkod", width=100, anchor="center")
        self.tree.column("Bolge", width=80, anchor="center")
        self.tree.column("Personel", width=120, anchor="center")
        self.tree.column("Sicil", width=80, anchor="center")
        self.tree.column("Tip", width=80, anchor="center")
        self.tree.column("Seri", width=100, anchor="center")
        self.tree.column("Gonderim", width=100, anchor="center")
        self.tree.column("Gelme", width=100, anchor="center")
        self.tree.column("Durum", width=80, anchor="center")
        self.tree.column("Aciklama", width=150, anchor="center")
        self.tree.column("Belge", width=80, anchor="center")
        self.tree.pack(fill="both", expand=True)

        style.configure("Treeview", background="#ffffff", fieldbackground="#ffffff", font=("Helvetica", 10))
        style.configure("Treeview.Heading", background="#2c3e50", foreground="white", font=("Helvetica", 11, "bold"))
        self.tree.tag_configure("all", background="#ffffff", foreground="#333333")

        scrollbar.config(command=self.tree.yview)

        style.configure("Success.TButton", background="#28a745")
        style.configure("Info.TButton", background="#007bff")
        style.configure("Warning.TButton", background="#ffc107")
        style.configure("Danger.TButton", background="#dc3545")
        style.configure("Primary.TButton", background="#6c757d")
        self.durum_renkleri = {
            "Serviste": "#ffcccc",
            "Servise Gönderildi": "#fff3cd",
            "Tamir edildi": "#d4edda",
            "Tamir olmuyor": "#cce5ff",
            "Hurda": "#cce5ff"
        }
        for durum, renk in self.durum_renkleri.items():
            self.tree.tag_configure(durum, background=renk)

        self.tree.bind("<Double-1>", self.on_tree_double_click)

        self.root.protocol("WM_DELETE_WINDOW", self.kapat)
        self.tum_cihazlari_listele()

    def show_settings(self):
        settings_win = tk.Toplevel(self.root)
        settings_win.title("Ayarlar")
        settings_win.geometry("400x200")
        settings_win.configure(bg="#f0f0f0")
        settings_win.resizable(False, False)

        settings_frame = ttk.LabelFrame(settings_win, text="Uygulama Ayarları", padding="10")
        settings_frame.pack(fill="x", pady=10)

        log_var = tk.BooleanVar(value=settings["log_enabled"])
        ttk.Checkbutton(settings_frame, text="Log Kayıtlarını Tut", 
                        variable=log_var, command=lambda: self.toggle_logging(log_var)).grid(row=0, column=0, columnspan=2, pady=5, sticky="w")

        ttk.Label(settings_frame, text="Log Dosyası Yolu:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.log_file_entry = ttk.Entry(settings_frame, width=30)
        self.log_file_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        self.log_file_entry.insert(0, settings["log_file"])
        ttk.Button(settings_frame, text="Dosya Seç", command=self.select_log_file).grid(row=1, column=2, padx=5, pady=5)

        ttk.Button(settings_frame, text="Kaydet", command=lambda: self.save_settings_from_ui(log_var)).grid(row=2, column=0, pady=10)
        ttk.Button(settings_frame, text="Kapat", command=settings_win.destroy).grid(row=2, column=1, pady=10)

    def toggle_logging(self, log_var):
        settings["log_enabled"] = log_var.get()
        configure_logging()
        if settings["log_enabled"]:
            messagebox.showinfo("Bilgi", f"Loglama etkinleştirildi. Kayıtlar '{settings['log_file']}' dosyasına yazılacak.")
        else:
            messagebox.showinfo("Bilgi", "Loglama devre dışı bırakıldı.")

    def select_log_file(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".log", 
                                                filetypes=[("Log Dosyaları", "*.log"), ("Tüm Dosyalar", "*.*")],
                                                initialfile=os.path.basename(settings["log_file"]))
        if file_path:
            self.log_file_entry.delete(0, tk.END)
            self.log_file_entry.insert(0, file_path)

    def save_settings_from_ui(self, log_var):
        settings["log_enabled"] = log_var.get()
        settings["log_file"] = self.log_file_entry.get()
        save_settings()
        configure_logging()
        messagebox.showinfo("Başarılı", "Ayarlar kaydedildi!")

    def on_tree_double_click(self, event):
        item = self.tree.selection()
        if item:
            values = self.tree.item(item[0], "values")
            id = values[0]
            self.load_cihaz_to_entries(id)
        else:
            messagebox.showwarning("Hata", "Lütfen bir cihaz seçin!")

    def load_cihaz_to_entries(self, id):
        try:
            cihaz = self.db.fetch_cihaz_by_id(id)
            if cihaz:
                self.entries["barkod_no"].delete(0, tk.END)
                self.entries["barkod_no"].insert(0, cihaz[1] or "")
                self.entries["bolge"].delete(0, tk.END)
                self.entries["bolge"].insert(0, cihaz[2] or "")
                self.entries["personel_ad_soyad"].delete(0, tk.END)
                self.entries["personel_ad_soyad"].insert(0, cihaz[3] or "")
                self.entries["personel_sicil_no"].delete(0, tk.END)
                self.entries["personel_sicil_no"].insert(0, cihaz[4] or "")
                self.entries["cihaz_tipi"].set(cihaz[5] or "")
                self.entries["cihaz_seri_no"].delete(0, tk.END)
                self.entries["cihaz_seri_no"].insert(0, cihaz[6] or "")
                self.entries["servis_gonderim_tarihi"].set_date(datetime.datetime.strptime(cihaz[7], "%d.%m.%Y") if cihaz[7] else datetime.datetime.now())
                self.entries["servis_gelme_tarihi"].set_date(datetime.datetime.strptime(cihaz[8], "%d.%m.%Y") if cihaz[8] else datetime.datetime.now())
                self.entries["cihaz_durumu"].set(cihaz[9] or "")
                self.entries["aciklama"].delete("1.0", tk.END)
                self.entries["aciklama"].insert("1.0", cihaz[10] or "")
                self.dosya_label.config(text=f"{len(cihaz[11])} belge yüklü" if cihaz[11] else "Henüz dosya seçilmedi")
                self.secilen_dosyalar = []
                self.selected_id = id
            else:
                messagebox.showwarning("Hata", f"ID {id} için cihaz bulunamadı!")
        except Exception as e:
            if settings["log_enabled"]:
                logging.error(f"Cihaz yükleme hatası (ID: {id}): {e}")
            messagebox.showerror("Hata", f"Cihaz bilgileri yüklenemedi: {e}")

    def check_aciklama_length(self, event):
        content = self.entries["aciklama"].get("1.0", tk.END).strip()
        if len(content) > 300:
            messagebox.showwarning("Uyarı", "Açıklama 300 karakteri geçemez!")
            self.entries["aciklama"].delete("1.0", tk.END)
            self.entries["aciklama"].insert("1.0", content[:300])

    def simple_search(self):
        filtre = self.search_var.get()
        self.tree.delete(*self.tree.get_children())
        for cihaz in self.db.fetch_all(filtre):
            item = self.tree.insert("", "end", values=(cihaz[0],) + cihaz[1:-1] + (len(cihaz[11]),))
            self.tree.item(item, tags=(cihaz[9],))

    def dosyalar_sec(self):
        dosyalar = filedialog.askopenfilenames(filetypes=[("Tüm Dosyalar", "*.*")])
        if dosyalar:
            self.secilen_dosyalar.extend(dosyalar)
            self.dosya_label.config(text=f"{len(self.secilen_dosyalar)} dosya seçildi")

    def show_belgeler(self):
        barkod_no = self.entries["barkod_no"].get()
        if not barkod_no:
            messagebox.showwarning("Hata", "Lütfen bir barkod numarası girin!")
            return

        cihazlar = self.db.fetch_cihazlar_by_barkod(barkod_no)
        if not cihazlar:
            messagebox.showinfo("Bilgi", "Bu barkod numarasına ait cihaz bulunamadı!")
            return

        belge_win = tk.Toplevel(self.root)
        belge_win.title(f"{barkod_no} - Yüklenen Belgeler")
        belge_win.geometry("600x400")
        belge_win.configure(bg="#f0f0f0")

        tree = ttk.Treeview(belge_win, columns=("ID", "Tarih", "Dosya"), show="headings")
        tree.heading("ID", text="Kayıt ID")
        tree.heading("Tarih", text="Gönderim Tarihi")
        tree.heading("Dosya", text="Dosya Adı")
        tree.column("ID", width=60)
        tree.column("Tarih", width=100)
        tree.column("Dosya", width=400)
        tree.pack(fill="both", expand=True, padx=10, pady=10)

        for cihaz in cihazlar:
            for belge in cihaz[11]:
                tree.insert("", "end", values=(cihaz[0], cihaz[7], os.path.basename(belge)))

        def open_belge():
            selected = tree.selection()
            if selected:
                id, _, dosya_adi = tree.item(selected[0])["values"]
                for cihaz in cihazlar:
                    if cihaz[0] == id:
                        for belge in cihaz[11]:
                            if os.path.basename(belge) == dosya_adi:
                                if os.path.exists(belge):
                                    os.startfile(belge)
                                else:
                                    messagebox.showwarning("Hata", "Belge bulunamadı!")
                                break
                        break

        def delete_belge():
            selected = tree.selection()
            if selected:
                id, _, dosya_adi = tree.item(selected[0])["values"]
                for cihaz in cihazlar:
                    if cihaz[0] == id:
                        belgeler = cihaz[11]
                        for i, belge in enumerate(belgeler):
                            if os.path.basename(belge) == dosya_adi:
                                if messagebox.askyesno("Onay", f"'{dosya_adi}' belgesi silinsin mi?"):
                                    belgeler.pop(i)
                                    yeni_veriler = {
                                        "barkod_no": cihaz[1], "bolge": cihaz[2], "personel_ad_soyad": cihaz[3],
                                        "personel_sicil_no": cihaz[4], "cihaz_tipi": cihaz[5], "cihaz_seri_no": cihaz[6],
                                        "servis_gonderim_tarihi": cihaz[7], "servis_gelme_tarihi": cihaz[8],
                                        "cihaz_durumu": cihaz[9], "aciklama": cihaz[10], "cihaz_belgeleri": belgeler
                                    }
                                    self.db.update_cihaz(yeni_veriler, id)
                                    tree.delete(selected[0])
                                    self.tum_cihazlari_listele()
                                break
                        break

        btn_frame = ttk.Frame(belge_win)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="Aç", command=open_belge).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Sil", command=delete_belge).pack(side="left", padx=5)

    def cihaz_kaydet(self):
        try:
            veriler = {key: entry.get() for key, entry in self.entries.items() if key != "aciklama"}
            veriler["aciklama"] = self.entries["aciklama"].get("1.0", tk.END).strip()
            if not veriler["barkod_no"]:
                messagebox.showwarning("Hata", "Barkod No zorunludur!")
                return
            if veriler["personel_sicil_no"] and not re.match(r"^\d{5,}$", veriler["personel_sicil_no"]):
                messagebox.showwarning("Hata", "Personel Sicil No en az 5 rakam olmalı!")
                return

            try:
                gonderim = datetime.datetime.strptime(veriler["servis_gonderim_tarihi"], "%d.%m.%Y")
                gelme = datetime.datetime.strptime(veriler["servis_gelme_tarihi"], "%d.%m.%Y")
                if gonderim > gelme:
                    messagebox.showwarning("Hata", "Gönderim tarihi, gelme tarihinden sonra olamaz!")
                    return
            except ValueError as e:
                if settings["log_enabled"]:
                    logging.error(f"Tarih formatı hatası: {e}")
                messagebox.showerror("Hata", f"Tarih formatı hatalı: {e}")
                return

            veriler["cihaz_belgeleri"] = []
            cihaz_id = self.db.insert_cihaz(veriler)
            if not cihaz_id:
                messagebox.showerror("Hata", "Cihaz veritabanına kaydedilemedi!")
                return

            belge_yollari = []
            if self.secilen_dosyalar:
                os.makedirs("belgeler", exist_ok=True)
                for dosya in self.secilen_dosyalar:
                    if not os.path.exists(dosya):
                        messagebox.showwarning("Hata", f"Dosya bulunamadı: {dosya}")
                        continue
                    barkod_no_clean = re.sub(r'[<>:"/\\|?*]', '_', veriler['barkod_no'])
                    seri_no_clean = re.sub(r'[<>:"/\\|?*]', '_', veriler['cihaz_seri_no'] or "bos")
                    dosya_adi_clean = re.sub(r'[<>:"/\\|?*]', '_', os.path.basename(dosya))
                    yeni_dosya_adi = f"{cihaz_id}_{barkod_no_clean}_{seri_no_clean}_{dosya_adi_clean}"
                    yeni_dosya_yolu = os.path.join("belgeler", yeni_dosya_adi)
                    try:
                        shutil.copy(dosya, yeni_dosya_yolu)
                        belge_yollari.append(yeni_dosya_yolu)
                    except Exception as e:
                        if settings["log_enabled"]:
                            logging.error(f"Dosya kopyalama hatası: {e} (Dosya: {dosya})")
                        messagebox.showwarning("Hata", f"Dosya kopyalanamadı: {dosya}")
                        continue

                if belge_yollari:
                    veriler["cihaz_belgeleri"] = belge_yollari
                    if not self.db.update_cihaz(veriler, cihaz_id):
                        messagebox.showerror("Hata", "Belgeler güncellenemedi!")
                        return

            messagebox.showinfo("Başarılı", "Cihaz başarıyla kaydedildi!")
            self.temizle()
            self.tum_cihazlari_listele()

        except Exception as e:
            if settings["log_enabled"]:
                logging.error(f"Cihaz kaydetme hatası: {e}")
            messagebox.showerror("Hata", f"Beklenmedik bir hata oluştu: {e}")

    def durum_guncelle(self):
        if not hasattr(self, "selected_id"):
            messagebox.showwarning("Hata", "Lütfen listeden bir cihaz seçin!")
            return

        try:
            yeni_veriler = {key: entry.get() for key, entry in self.entries.items() if key != "aciklama"}
            yeni_veriler["aciklama"] = self.entries["aciklama"].get("1.0", tk.END).strip()
            if yeni_veriler["personel_sicil_no"] and not re.match(r"^\d{5,}$", yeni_veriler["personel_sicil_no"]):
                messagebox.showwarning("Hata", "Personel Sicil No en az 5 rakam olmalı!")
                return

            gonderim = datetime.datetime.strptime(yeni_veriler["servis_gonderim_tarihi"], "%d.%m.%Y")
            gelme = datetime.datetime.strptime(yeni_veriler["servis_gelme_tarihi"], "%d.%m.%Y")
            if gonderim > gelme:
                messagebox.showwarning("Hata", "Gönderim tarihi, gelme tarihinden sonra olamaz!")
                return

            mevcut_cihaz = self.db.fetch_cihaz_by_id(self.selected_id)
            belge_yollari = mevcut_cihaz[11].copy()  # Mevcut belgelerin bir kopyasını al

            # Yeni belgeleri ekle
            if self.secilen_dosyalar:
                os.makedirs("belgeler", exist_ok=True)
                for dosya in self.secilen_dosyalar:
                    if not os.path.exists(dosya):
                        messagebox.showwarning("Hata", f"Dosya bulunamadı: {dosya}")
                        continue
                    barkod_no_clean = re.sub(r'[<>:"/\\|?*]', '_', yeni_veriler['barkod_no'])
                    seri_no_clean = re.sub(r'[<>:"/\\|?*]', '_', yeni_veriler['cihaz_seri_no'] or "bos")
                    dosya_adi_clean = re.sub(r'[<>:"/\\|?*]', '_', os.path.basename(dosya))
                    yeni_dosya_adi = f"{self.selected_id}_{barkod_no_clean}_{seri_no_clean}_{dosya_adi_clean}"
                    yeni_dosya_yolu = os.path.join("belgeler", yeni_dosya_adi)
                    try:
                        shutil.copy(dosya, yeni_dosya_yolu)
                        if yeni_dosya_yolu not in belge_yollari:  # Aynı dosyanın tekrar eklenmesini önle
                            belge_yollari.append(yeni_dosya_yolu)
                        if settings["log_enabled"]:
                            logging.info(f"Yeni belge eklendi: {yeni_dosya_yolu}")
                    except Exception as e:
                        if settings["log_enabled"]:
                            logging.error(f"Dosya kopyalama hatası: {e} (Dosya: {dosya})")
                        messagebox.showwarning("Hata", f"Dosya kopyalanamadı: {dosya}")
                        continue

            yeni_veriler["cihaz_belgeleri"] = belge_yollari  # Güncellenmiş belgeleri ata

            degisiklikler = []
            alanlar_ve_indeksler = {
                "bolge": 2,
                "personel_ad_soyad": 3,
                "personel_sicil_no": 4,
                "cihaz_tipi": 5,
                "barkod_no": 1,
                "cihaz_seri_no": 6,
                "servis_gonderim_tarihi": 7,
                "servis_gelme_tarihi": 8,
                "cihaz_durumu": 9,
                "aciklama": 10
            }
            
            for key, indeks in alanlar_ve_indeksler.items():
                eski_deger = str(mevcut_cihaz[indeks] or "").strip()
                yeni_deger = str(yeni_veriler.get(key, "") or "").strip()
                if eski_deger != yeni_deger and (eski_deger or yeni_deger):
                    degisiklikler.append((key.replace('_', ' ').title(), eski_deger, yeni_deger))

            if len(yeni_veriler["cihaz_belgeleri"]) > len(mevcut_cihaz[11]):
                degisiklikler.append(("Cihaz Belgeleri", str(len(mevcut_cihaz[11])), str(len(yeni_veriler["cihaz_belgeleri"]))))

            if not degisiklikler:
                messagebox.showinfo("Bilgi", "Herhangi bir değişiklik yapılmadı.")
                return

        except ValueError as e:
            messagebox.showerror("Hata", f"Tarih formatı hatalı: {e}")
            return
        except Exception as e:
            if settings["log_enabled"]:
                logging.error(f"Güncelleme hatası: {e}")
            messagebox.showerror("Hata", f"Beklenmedik bir hata oluştu: {e}")
            return

        onay_win = tk.Toplevel(self.root)
        onay_win.title("Güncelleme Onayı")
        onay_win.geometry("500x400")
        onay_win.configure(bg="#ffffff")
        onay_win.resizable(False, False)
        onay_win.transient(self.root)
        onay_win.grab_set()

        ttk.Label(onay_win, text=f"Güncelleme Onayı (Kayıt ID: {self.selected_id})", 
                  font=("Helvetica", 14, "bold"), foreground="#2c3e50").pack(pady=15, padx=20)

        style = ttk.Style()
        style.configure("Custom.Treeview", background="#f8f9fa", fieldbackground="#f8f9fa", font=("Arial", 10))
        style.configure("Custom.Treeview.Heading", background="#2c3e50", foreground="white", font=("Helvetica", 11, "bold"))

        tree = ttk.Treeview(onay_win, columns=("Alan", "Eski Değer", "Yeni Değer"), show="headings", 
                            style="Custom.Treeview", height=len(degisiklikler) + 1)
        tree.heading("Alan", text="Alan")
        tree.heading("Eski Değer", text="Eski Değer")
        tree.heading("Yeni Değer", text="Yeni Değer")
        tree.column("Alan", width=150, anchor="w")
        tree.column("Eski Değer", width=150, anchor="w")
        tree.column("Yeni Değer", width=150, anchor="w")
        tree.pack(fill="both", expand=True, padx=20, pady=10)

        for alan, eski, yeni in degisiklikler:
            tree.insert("", "end", values=(alan, eski, yeni))

        style.configure("Confirm.TButton", font=("Helvetica", 10, "bold"), 
                       background="#4CAF50", foreground="white", padding=8)
        style.map("Confirm.TButton", background=[("active", "#45a049")])

        style.configure("Cancel.TButton", font=("Helvetica", 10, "bold"), 
                       background="#f44336", foreground="white", padding=8)
        style.map("Cancel.TButton", background=[("active", "#da190b")])

        style.configure("ButtonFrame.TFrame", background="#ffffff")
        btn_frame = ttk.Frame(onay_win, style="ButtonFrame.TFrame")
        btn_frame.pack(pady=15, padx=20)

        ttk.Button(btn_frame, text="Onayla", command=lambda: self.onay_kapat(onay_win, yeni_veriler, degisiklikler, True), 
                   style="Confirm.TButton").pack(side="left", padx=10)
        ttk.Button(btn_frame, text="İptal", command=lambda: self.onay_kapat(onay_win, yeni_veriler, degisiklikler, False), 
                   style="Cancel.TButton").pack(side="left", padx=10)

    def onay_kapat(self, window, veriler, degisiklikler, onay):
        window.destroy()
        if onay and self.db.update_cihaz(veriler, self.selected_id):
            if settings["log_enabled"]:
                logging.info(f"Cihaz güncellendi: ID {self.selected_id} | Değişiklikler: {', '.join([f'{d[0]}: {d[1]} -> {d[2]}' for d in degisiklikler])}")
            messagebox.showinfo("Başarılı", f"ID {self.selected_id} başarıyla güncellendi!\n\nDeğişiklikler:\n" + "\n".join([f"{d[0]}: '{d[1]}' -> '{d[2]}'" for d in degisiklikler]))
            self.tum_cihazlari_listele()
            self.temizle()
            delattr(self, "selected_id")
        elif onay:
            messagebox.showerror("Hata", "Güncelleme sırasında bir sorun oluştu!")

    def cihaz_sil(self):
        if not hasattr(self, "selected_id"):
            messagebox.showwarning("Hata", "Lütfen listeden bir cihaz seçin!")
            return

        if messagebox.askyesno("Onay", f"ID {self.selected_id} numaralı kayıt silinsin mi?"):
            self.db.delete_cihaz(self.selected_id)
            messagebox.showinfo("Başarılı", "Kayıt silindi!")
            self.temizle()
            self.tum_cihazlari_listele()
            delattr(self, "selected_id")

    def tum_cihazlari_listele(self):
        self.temizle()
        self.tree.delete(*self.tree.get_children())
        for cihaz in self.db.fetch_all():
            item = self.tree.insert("", "end", values=(cihaz[0],) + cihaz[1:-1] + (len(cihaz[11]),))
            self.tree.item(item, tags=(cihaz[9],))

    def show_advanced_search(self):
        search_win = tk.Toplevel(self.root)
        search_win.title("Sorgulama")
        search_win.geometry("300x150")
        search_win.configure(bg="#f0f0f0")
        search_win.resizable(False, False)

        filter_frame = ttk.LabelFrame(search_win, text="Filtreler", padding="10")
        filter_frame.pack(fill="x", pady=10)

        ttk.Label(filter_frame, text="Barkod No:").grid(row=0, column=0, padx=5, pady=5)
        barkod_entry = ttk.Entry(filter_frame)
        barkod_entry.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(filter_frame, text="Cihaz Durumu:").grid(row=1, column=0, padx=5, pady=5)
        durum_combo = ttk.Combobox(filter_frame, values=["", "Serviste", "Servise Gönderildi", "Tamir edildi","Tamir olmuyor","Hurda"])
        durum_combo.grid(row=1, column=1, padx=5, pady=5)

        def perform_search():
            filtreler = {
                "barkod_no": barkod_entry.get(),
                "cihaz_durumu": durum_combo.get() if durum_combo.get() else None
            }
            sonuclar = self.db.advanced_search(filtreler)
            self.tree.delete(*self.tree.get_children())
            for cihaz in sonuclar:
                item = self.tree.insert("", "end", values=(cihaz[0],) + cihaz[1:-1] + (len(cihaz[11]),))
                self.tree.item(item, tags=(cihaz[9],))
            search_win.destroy()

        ttk.Button(filter_frame, text="Sorgula", command=perform_search).grid(row=2, columnspan=2, pady=10)

    def export_to_excel(self):
        veriler = [self.tree.item(item)["values"] for item in self.tree.get_children()]
        if not veriler:
            messagebox.showinfo("Bilgi", "Dışa aktarılacak veri bulunamadı!")
            return

        df = pd.DataFrame(veriler, columns=["Kayıt ID", "Barkod No", "Bölge", "Personel Ad Soyad", "Personel Sicil No",
                                            "Cihaz Tipi", "Cihaz Seri No", "Servis Gönderim Tarihi",
                                            "Servis Gelme Tarihi", "Cihaz Durumu", "Açıklama", "Belge Sayısı"])
        dosya_adi = f"servis_takip_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        try:
            df.to_excel(dosya_adi, index=False)
            messagebox.showinfo("Başarılı", f"Veriler {dosya_adi} dosyasına aktarıldı!")
            if settings["log_enabled"]:
                logging.info(f"Excel'e aktarma yapıldı: {dosya_adi}")
        except Exception as e:
            messagebox.showerror("Hata", f"Excel'e aktarma başarısız: {e}")

    def temizle(self):
        for key, entry in self.entries.items():
            if key in ["servis_gonderim_tarihi", "servis_gelme_tarihi"]:
                entry.set_date(datetime.datetime.now())
            elif key in ["cihaz_tipi", "cihaz_durumu"]:
                entry.set(entry["values"][0])
            elif key == "aciklama":
                entry.delete("1.0", tk.END)
            else:
                entry.delete(0, tk.END)
        self.dosya_label.config(text="Henüz dosya seçilmedi")
        self.secilen_dosyalar = []
        if hasattr(self, "selected_id"):
            delattr(self, "selected_id")
        self.search_var.set("")

    def kapat(self):
        self.db.close()
        self.root.destroy()

if __name__ == "__main__":
    app = ServisTakipUygulamasi()
    app.root.mainloop()