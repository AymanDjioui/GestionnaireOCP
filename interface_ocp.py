import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
import pandas as pd
from PIL import Image, ImageTk
import os
import shutil
from datetime import datetime
import threading
from concurrent.futures import ThreadPoolExecutor
import math
import hashlib
import stat
import sys

def resource_path(relative_path):
    # Trouve le bon chemin pour PyInstaller ou pour le script normal
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

class DatabaseManager:
    """Gestionnaire de base de données SQLite"""
    def __init__(self, db_path="ocp_pieces.db"):
        self.db_path = db_path
        self.init_database()

    def init_database(self):
        """Initialiser la base de données"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS pieces (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                article TEXT NOT NULL,
                code_sap TEXT,
                description TEXT,
                description_longue TEXT,
                unite_mesure TEXT,
                statut_article TEXT,
                quantite_installee TEXT,
                situation TEXT,
                image_path TEXT,
                date_creation TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                date_modification TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        # Ajout dynamique des colonnes si elles n'existent pas déjà
        try:
            cursor.execute("ALTER TABLE pieces ADD COLUMN quantite_installee TEXT")
        except sqlite3.OperationalError:
            pass
        try:
            cursor.execute("ALTER TABLE pieces ADD COLUMN situation TEXT")
        except sqlite3.OperationalError:
            pass
        # Index pour optimiser les recherches
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_article ON pieces(article)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_code_sap ON pieces(code_sap)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_description ON pieces(description)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_statut ON pieces(statut_article)')
        conn.commit()
        conn.close()

    def migrate_from_excel(self, excel_path):
        """Migrer les données depuis Excel vers SQLite"""
        if not os.path.exists(excel_path):
            return False
        try:
            df = pd.read_excel(excel_path)
            conn = sqlite3.connect(self.db_path)
            # Remplacer les NaN par des chaînes vides pour tous les champs
            df = df.fillna("")
            cursor = conn.cursor()
            cursor.execute("SELECT COUNT(*) FROM pieces")
            count = cursor.fetchone()[0]
            if count == 0:
                for _, row in df.iterrows():
                    cursor.execute('''
                        INSERT INTO pieces
                        (article, code_sap, description, description_longue, unite_mesure, statut_article, quantite_installee, situation, image_path)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (
                        str(row.get("Article", "")),
                        str(row.get("code SAP", "")),
                        str(row.get("Description", "")),
                        str(row.get("Description longue", "")),
                        str(row.get("Unité de mesure principale", "")),
                        str(row.get("Statut de l'article", "")),
                        str(row.get("Quantité installée", "")),
                        str(row.get("Situation", "")),
                        str(row.get("Image", ""))
                    ))
                conn.commit()
                print(f"Migration terminée: {len(df)} enregistrements importés")
            conn.close()
            return True
        except Exception as e:
            print(f"Erreur migration: {e}")
            return False

    def search_pieces(self, filters=None, limit=1000, offset=0):
        """Rechercher des pièces avec filtres"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        query = "SELECT * FROM pieces WHERE 1=1"
        params = []
        if filters:
            if filters.get('article'):
                query += " AND article LIKE ?"
                params.append(f"%{filters['article']}%")
            # Ajout du filtre pour code SAP vide
            if filters.get('code_sap_empty'):
                query += " AND (code_sap IS NULL OR code_sap='' OR lower(code_sap)='nan')"
            elif filters.get('code_sap'):
                query += " AND code_sap LIKE ?"
                params.append(f"%{filters['code_sap']}%")
            if filters.get('description'):
                query += " AND description LIKE ?"
                params.append(f"%{filters['description']}%")
            if filters.get('description_longue'):
                query += " AND description_longue LIKE ?"
                params.append(f"%{filters['description_longue']}%")
            if filters.get('statut') and filters['statut'] != 'Tous':
                query += " AND statut_article LIKE ?"
                params.append(f"%{filters['statut']}%")
            if filters.get('unite') and filters['unite'] != 'Tous':
                query += " AND unite_mesure LIKE ?"
                params.append(f"%{filters['unite']}%")
            if filters.get('quantite_installee'):
                query += " AND quantite_installee LIKE ?"
                params.append(f"%{filters['quantite_installee']}%")
            if filters.get('situation'):
                query += " AND situation LIKE ?"
                params.append(f"%{filters['situation']}%")
        # Obtenir le nombre total d'abord
        count_query = query.replace("SELECT *", "SELECT COUNT(*)").split(" ORDER BY")[0]
        cursor.execute(count_query, params)
        total_count = cursor.fetchone()[0]
        # Ajouter le tri et la pagination pour la requête principale
        query += " ORDER BY article LIMIT ? OFFSET ?"
        params.extend([limit, offset])
        cursor.execute(query, params)
        results = cursor.fetchall()
        conn.close()
        return results, total_count

    def get_piece_by_id(self, piece_id):
        """Obtenir une pièce par ID"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM pieces WHERE id = ?", (piece_id,))
        result = cursor.fetchone()
        conn.close()
        return result

    def insert_piece(self, piece_data):
        """Insérer une nouvelle pièce"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO pieces
            (article, code_sap, description, description_longue, unite_mesure, statut_article, quantite_installee, situation, image_path)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', piece_data)
        piece_id = cursor.lastrowid
        conn.commit()
        conn.close()
        return piece_id

    def update_piece(self, piece_id, piece_data):
        """Mettre à jour une pièce"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('''
            UPDATE pieces
            SET article=?, code_sap=?, description=?, description_longue=?,
                unite_mesure=?, statut_article=?, quantite_installee=?, situation=?, image_path=?, date_modification=CURRENT_TIMESTAMP
            WHERE id=?
        ''', piece_data + (piece_id,))
        conn.commit()
        conn.close()

    def delete_piece(self, piece_id):
        """Supprimer une pièce"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute("DELETE FROM pieces WHERE id = ?", (piece_id,))
        conn.commit()
        conn.close()

    def export_to_excel(self, output_path, filters=None):
        """Exporter vers Excel"""
        conn = sqlite3.connect(self.db_path)
        query = "SELECT * FROM pieces WHERE 1=1"
        params = []
        if filters:
            if filters.get('article'):
                query += " AND article LIKE ?"
                params.append(f"%{filters['article']}%")
            if filters.get('code_sap'):
                query += " AND code_sap LIKE ?"
                params.append(f"%{filters['code_sap']}%")
            if filters.get('description'):
                query += " AND description LIKE ?"
                params.append(f"%{filters['description']}%")
            if filters.get('statut') and filters['statut'] != 'Tous':
                query += " AND statut_article LIKE ?"
                params.append(f"%{filters['statut']}%")
            if filters.get('unite') and filters['unite'] != 'Tous':
                query += " AND unite_mesure LIKE ?"
                params.append(f"%{filters['unite']}%")
            if filters.get('quantite_installee'):
                query += " AND quantite_installee LIKE ?"
                params.append(f"%{filters['quantite_installee']}%")
            if filters.get('situation'):
                query += " AND situation LIKE ?"
                params.append(f"%{filters['situation']}%")
        df = pd.read_sql_query(query, conn, params=params)
        df.rename(columns={
            'article': 'Article', 'code_sap': 'code SAP', 'description': 'Description',
            'description_longue': 'Description longue', 'unite_mesure': 'Unité de mesure principale',
            'statut_article': "Statut de l'article", 'quantite_installee': 'Quantité installée', 'situation': 'Situation', 'image_path': 'Image'
        }, inplace=True)
        df.drop(columns=['id', 'date_creation', 'date_modification'], inplace=True, errors='ignore')
        df.to_excel(output_path, index=False)
        conn.close()
        return len(df)

class OCPPiecesManager:
    HISTORIQUE_FILE = "historique.txt"
    MIGRATION_FLAG_FILE = "migration_done.flag"
    PASSWORD_FILE = "password.hash"
    def __init__(self, root):
        self.root = root
        self.root.title("Gestionnaire de Pièces OCP")
        self.root.configure(bg='#f8fafc')

        # Configuration responsive
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        window_width = min(max(int(screen_width * 0.85), 900), 1400)
        window_height = min(max(int(screen_height * 0.85), 650), 1000)
        
        self.root.geometry(f"{window_width}x{window_height}")
        self.root.minsize(900, 650)

        self.db_manager = DatabaseManager()
        self.current_page = 0
        self.page_size = 100
        self.total_records = 0
        self.current_piece_id = None
        self.current_image = None
        self.images_folder = "images_pieces"
        self.executor = ThreadPoolExecutor(max_workers=2)
        self.editing_mode = False

        if not os.path.exists(self.images_folder):
            os.makedirs(self.images_folder)
        self.check_and_run_migration()
        # Système de mot de passe : demander à chaque démarrage
        if not self.check_password():
            self.root.destroy()
            return
        # Suppression de la vérification d'intégrité
        self.setup_styles()
        self.create_widgets()
        
        self.setup_keyboard_shortcuts()
        self.create_help_menu()
        self.create_history_file_if_needed()

        self.load_data()
        self.update_button_states()

    def create_history_file_if_needed(self):
        if not os.path.exists(self.HISTORIQUE_FILE):
            with open(self.HISTORIQUE_FILE, 'w', encoding='utf-8') as f:
                f.write("Historique des modifications\n===========================\n")

    def log_history(self, action, piece_id=None, details=None, old_data=None, new_data=None):
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        entry = f"[{now}] Action: {action}"
        if piece_id is not None:
            entry += f" | ID: {piece_id}"
        if details:
            entry += f" | {details}"
        # Ajout des détails avancés
        if old_data is not None or new_data is not None:
            entry += "\n"
            if old_data is not None and new_data is not None:
                # Modification : afficher les champs modifiés
                for k in new_data.keys():
                    old_val = old_data.get(k, "") if old_data else ""
                    new_val = new_data.get(k, "") if new_data else ""
                    if str(old_val) != str(new_val):
                        entry += f"    {k} : '{old_val}' -> '{new_val}'\n"
            elif new_data is not None:
                # Création : afficher toutes les valeurs
                for k, v in new_data.items():
                    entry += f"    {k} : '{v}'\n"
            elif old_data is not None:
                # Suppression : afficher toutes les anciennes valeurs
                for k, v in old_data.items():
                    entry += f"    {k} : '{v}'\n"
        entry += "\n"
        with open(self.HISTORIQUE_FILE, 'a', encoding='utf-8') as f:
            f.write(entry)

    def hash_password(self, password):
        return hashlib.sha256(password.encode('utf-8')).hexdigest()

    def set_password(self):
        # Fenêtre de création de mot de passe (2 fois)
        win = tk.Toplevel(self.root)
        win.title("Créer un mot de passe")
        win.geometry("350x240")  # Agrandie pour le bouton
        win.resizable(False, False)
        win.transient(self.root)
        win.grab_set()
        frame = ttk.Frame(win, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frame, text="Créer un mot de passe d'accès", font=("Segoe UI", 12, "bold")).pack(pady=(0, 10))
        ttk.Label(frame, text="Mot de passe :").pack(anchor=tk.W)
        pwd1 = ttk.Entry(frame, show="*", width=25)
        pwd1.pack(pady=(0, 8))
        ttk.Label(frame, text="Confirmer le mot de passe :").pack(anchor=tk.W)
        pwd2 = ttk.Entry(frame, show="*", width=25)
        pwd2.pack(pady=(0, 8))
        msg = ttk.Label(frame, text="", foreground="#ef4444")
        msg.pack()
        def save_pwd():
            p1, p2 = pwd1.get(), pwd2.get()
            if not p1 or not p2:
                msg.config(text="Veuillez remplir les deux champs.")
                return
            if p1 != p2:
                msg.config(text="Les mots de passe ne correspondent pas.")
                return
            with open(self.PASSWORD_FILE, 'w') as f:
                f.write(self.hash_password(p1))
            win.destroy()
        ttk.Button(frame, text="Valider", command=save_pwd, style="Success.TButton").pack(pady=(18, 0))
        self.root.wait_window(win)

    def check_password(self):
        # Si le fichier n'existe pas, demander la création
        if not os.path.exists(self.PASSWORD_FILE):
            self.set_password()
        # Demander le mot de passe à chaque démarrage
        for attempt in range(3):
            win = tk.Toplevel(self.root)
            win.title("Sécurité - Mot de passe")
            win.geometry("350x160")
            win.resizable(False, False)
            win.transient(self.root)
            win.grab_set()
            frame = ttk.Frame(win, padding="20")
            frame.pack(fill=tk.BOTH, expand=True)
            ttk.Label(frame, text="Veuillez saisir votre mot de passe", font=("Segoe UI", 12, "bold")).pack(pady=(0, 10))
            pwd = ttk.Entry(frame, show="*", width=25)
            pwd.pack(pady=(0, 8))
            msg = ttk.Label(frame, text="", foreground="#ef4444")
            msg.pack()
            result = {'ok': False, 'closed': False}
            def check():
                with open(self.PASSWORD_FILE, 'r') as f:
                    hash_saved = f.read().strip()
                if self.hash_password(pwd.get()) == hash_saved:
                    result['ok'] = True
                win.destroy()
            def on_close():
                result['closed'] = True
                win.destroy()
            win.protocol("WM_DELETE_WINDOW", on_close)
            ttk.Button(frame, text="Valider", command=check, style="Primary.TButton").pack(pady=(10, 0))
            self.root.wait_window(win)
            if result['ok']:
                return True
            if result['closed']:
                self.root.after(100, self.root.destroy)
                return False
        messagebox.showerror("Sécurité", "3 tentatives échouées. L'application va se fermer.")
        self.root.after(100, self.root.destroy)
        return False

    def check_and_run_migration(self):
        # N'afficher la migration que si le flag n'existe pas
        if not os.path.exists(self.MIGRATION_FLAG_FILE):
            self.migrate_excel_data()
            # Créer le flag après la migration ou refus
            with open(self.MIGRATION_FLAG_FILE, 'w') as f:
                f.write('done')
            # Suppression de la protection des fichiers sensibles et du hash après migration

    def migrate_excel_data(self):
        if os.path.exists("data.xlsx"):
            if messagebox.askyesno("Migration",
                                 "Fichier Excel détecté. Voulez-vous migrer les données vers SQLite?\n"
                                 "Cette opération ne sera effectuée qu'une seule fois."):
                if self.db_manager.migrate_from_excel("data.xlsx"):
                    messagebox.showinfo("Migration", "Migration terminée avec succès!")
                else:
                    messagebox.showerror("Migration", "Erreur lors de la migration")
            # Après la migration, demander la création du mot de passe si besoin
            if not os.path.exists(self.PASSWORD_FILE):
                self.set_password()

    def setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        colors = {
            'primary': '#2563eb', 'secondary': '#64748b', 'success': '#10b981',
            'danger': '#ef4444', 'warning': '#f59e0b', 'light': '#f8fafc',
            'dark': '#1e293b', 'white': '#ffffff', 'border': '#e2e8f0',
            'row_alt': '#f1f5f9', 'focus': '#dbeafe', 'hover': '#e0e7ff'
        }

        # Moderniser les boutons
        style.configure("Action.TButton", font=("Segoe UI", 10, "bold"), relief="flat", borderwidth=0, focuscolor="none", padding=6, border=0)
        style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"), relief="flat", borderwidth=0, background=colors['primary'], foreground=colors['white'], focuscolor="none", padding=6, border=0)
        style.map("Primary.TButton",
            background=[('active', colors['focus']), ('!active', colors['primary'])],
            foreground=[('active', colors['primary']), ('!active', colors['white'])]
        )
        style.configure("Success.TButton", font=("Segoe UI", 10, "bold"), relief="flat", borderwidth=0, background=colors['success'], foreground=colors['white'], focuscolor="none", padding=6, border=0)
        style.configure("Danger.TButton", font=("Segoe UI", 10, "bold"), relief="flat", borderwidth=0, background=colors['danger'], foreground=colors['white'], focuscolor="none", padding=6, border=0)
        style.configure("Search.TButton", font=("Segoe UI", 9), relief="flat", borderwidth=0, focuscolor="none", padding=5, border=0)
        style.configure("Blue.TButton", font=("Segoe UI", 9, "bold"), relief="flat", borderwidth=0, focuscolor="none", background=colors['primary'], foreground=colors['white'], padding=5, border=0)
        style.configure("Red.TButton", font=("Segoe UI", 9, "bold"), relief="flat", borderwidth=0, focuscolor="none", background=colors['danger'], foreground=colors['white'], padding=5, border=0)
        style.configure("Green.TButton", font=("Segoe UI", 9, "bold"), relief="flat", borderwidth=0, focuscolor="none", background=colors['success'], foreground=colors['white'], padding=5, border=0)

        # Moderniser les labelframes et entrées
        style.configure("Modern.TLabelframe", borderwidth=1, relief="solid", background=colors['white'], padding=10)
        style.configure("Modern.TLabelframe.Label", font=("Segoe UI", 11, "bold"), foreground=colors['dark'])
        style.configure("Modern.TEntry", fieldbackground=colors['white'], borderwidth=1, relief="solid", insertcolor=colors['primary'], font=("Segoe UI", 10), padding=4)
        style.configure("Modern.TCombobox", fieldbackground=colors['white'], borderwidth=1, relief="solid", font=("Segoe UI", 10), padding=4)

        # Moderniser le Treeview
        style.configure("Modern.Treeview", background=colors['white'], foreground=colors['dark'], fieldbackground=colors['white'], borderwidth=0, relief="flat", font=("Segoe UI", 10))
        style.configure("Modern.Treeview.Heading", background=colors['light'], foreground=colors['dark'], font=("Segoe UI", 10, "bold"), relief="flat", borderwidth=1)
        style.map("Modern.Treeview",
            background=[('selected', colors['focus']), ('!selected', colors['white'])],
            foreground=[('selected', colors['primary']), ('!selected', colors['dark'])]
        )
        style.layout("Modern.Treeview", [
            ('Treeview.field', {'sticky': 'nswe', 'children': [
                ('Treeview.padding', {'sticky': 'nswe', 'children': [
                    ('Treeview.treearea', {'sticky': 'nswe'})
                ]})
            ]})
        ])

        # Alternance de lignes (row stripes)
        self.treeview_row_colors = (colors['white'], colors['row_alt'])

        style.configure("Modern.Horizontal.TProgressbar", background=colors['primary'], troughcolor=colors['border'], borderwidth=0, lightcolor=colors['primary'], darkcolor=colors['primary'])

    def create_widgets(self):
        style = ttk.Style()
        style.configure("Modern.TFrame", background='#f8fafc')
        main_frame = ttk.Frame(self.root, padding="18 12 18 12", style="Modern.TFrame")
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1)

        title_frame = ttk.Frame(main_frame, style="Modern.TFrame")
        title_frame.grid(row=0, column=0, sticky="ew", pady=(0, 24))
        title_frame.columnconfigure(0, weight=1)
        title_frame.columnconfigure(1, weight=1)
        title_frame.columnconfigure(2, weight=1)
        ttk.Label(title_frame, text="🔧 Gestionnaire de Pièces BDM - OCP KHOURIBGA", font=("Segoe UI", 22, "bold"), foreground="#1e293b", background="#f8fafc").grid(row=0, column=0, sticky="w", padx=(0, 10))
        ttk.Label(title_frame, text="MIK/GE/E - 217", font=("Segoe UI", 13), foreground="#64748b", background="#f8fafc").grid(row=0, column=1, sticky="w", padx=(15, 0))
        try:
            logo_path = resource_path("OCP_Group.svg.png")
            logo_image = Image.open(logo_path)
            logo_image.thumbnail((200, 100), Image.Resampling.LANCZOS)
            self.ocp_logo = ImageTk.PhotoImage(logo_image)
            logo_label = ttk.Label(title_frame, image=self.ocp_logo, background="#f8fafc")
            logo_label.grid(row=0, column=2, sticky="e", padx=10, pady=5)
        except FileNotFoundError:
            logo_label = ttk.Label(title_frame, text="Logo OCP", font=("Segoe UI", 10), background="#f8fafc")
            logo_label.grid(row=0, column=2, sticky="e", padx=10, pady=5)
        except Exception as e:
            print(f"Une erreur est survenue lors du chargement du logo : {e}")

        search_frame = ttk.LabelFrame(main_frame, text="🔍 Recherche et Filtres", padding="18 12 18 12", style="Modern.TLabelframe")
        search_frame.grid(row=1, column=0, sticky="ew", pady=(0, 18))
        search_frame.columnconfigure(0, weight=1)
        search_row1 = ttk.Frame(search_frame)
        search_row1.grid(row=0, column=0, sticky="ew", pady=(0, 7))
        search_row1.columnconfigure(tuple(range(10)), weight=1)
        ttk.Label(search_row1, text="Article:", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="w")
        self.search_article = ttk.Entry(search_row1, width=18, style="Modern.TEntry")
        self.search_article.grid(row=0, column=1, sticky="ew", padx=(5, 12))
        ttk.Label(search_row1, text="Code SAP:", font=("Segoe UI", 10, "bold")).grid(row=0, column=2, sticky="w")
        self.search_sap = ttk.Entry(search_row1, width=18, style="Modern.TEntry")
        self.search_sap.grid(row=0, column=3, sticky="ew", padx=(5, 12))
        ttk.Label(search_row1, text="Description:", font=("Segoe UI", 10, "bold")).grid(row=0, column=4, sticky="w")
        self.search_description = ttk.Entry(search_row1, width=22, style="Modern.TEntry")
        self.search_description.grid(row=0, column=5, sticky="ew", padx=(5, 12))
        ttk.Label(search_row1, text="Description longue:", font=("Segoe UI", 10, "bold")).grid(row=0, column=6, sticky="w")
        self.search_description_longue = ttk.Entry(search_row1, width=22, style="Modern.TEntry")
        self.search_description_longue.grid(row=0, column=7, sticky="ew", padx=(5, 12))

        search_row2 = ttk.Frame(search_frame)
        search_row2.grid(row=1, column=0, sticky="ew", pady=(0, 7))
        search_row2.columnconfigure(tuple(range(10)), weight=1)
        ttk.Label(search_row2, text="Statut:", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="w")
        self.search_statut = ttk.Combobox(search_row2, width=13, values=["Tous", "Actif", "Désactivé"], style="Modern.TCombobox")
        self.search_statut.grid(row=0, column=1, sticky="ew", padx=(5, 12))
        self.search_statut.set("Tous")
        ttk.Label(search_row2, text="Unité:", font=("Segoe UI", 10, "bold")).grid(row=0, column=2, sticky="w")
        unites_list = [
            "Tous", "LITRE", "KILOGRAMME", "TONNE", "PIECE", "MILLILITRE", "GRAMME", 
            "MÈTRE", "MÈTRE CARRÉ", "MÈTRE CUBE", "BARIL", "CENTIMÈTRE", "MILLIMÈTRE", 
            "UNITÉ", "BAR"
        ]
        self.search_unite = ttk.Combobox(search_row2, width=16, values=unites_list, style="Modern.TCombobox")
        self.search_unite.grid(row=0, column=3, sticky="ew", padx=(5, 12))
        self.search_unite.set("Tous")
        ttk.Label(search_row2, text="Qté installée:", font=("Segoe UI", 10, "bold")).grid(row=0, column=4, sticky="w")
        self.search_quantite_installee = ttk.Entry(search_row2, width=12, style="Modern.TEntry")
        self.search_quantite_installee.grid(row=0, column=5, sticky="ew", padx=(5, 12))
        ttk.Label(search_row2, text="Situation:", font=("Segoe UI", 10, "bold")).grid(row=0, column=6, sticky="w")
        self.search_situation = ttk.Entry(search_row2, width=14, style="Modern.TEntry")
        self.search_situation.grid(row=0, column=7, sticky="ew", padx=(5, 12))

        search_buttons = ttk.Frame(search_frame)
        search_buttons.grid(row=2, column=0, sticky="ew", pady=(7, 0))
        search_buttons.columnconfigure(tuple(range(4)), weight=1)
        btn_width = 15
        btn_rechercher = ttk.Button(search_buttons, text="🔎 Rechercher", command=self.search_data, style="Blue.TButton", width=btn_width)
        btn_rechercher.grid(row=0, column=0, sticky="ew", padx=(0, 5))
        btn_rechercher.tooltip = self.create_tooltip(btn_rechercher, "Lancer la recherche avec les filtres ci-dessus")
        btn_reset = ttk.Button(search_buttons, text="♻️ Réinitialiser", command=self.reset_search, style="Red.TButton", width=btn_width)
        btn_reset.grid(row=0, column=1, sticky="ew", padx=(0, 5))
        btn_reset.tooltip = self.create_tooltip(btn_reset, "Effacer tous les filtres de recherche")
        btn_export = ttk.Button(search_buttons, text="⬇️ Exporter Excel", command=self.export_to_excel, style="Green.TButton", width=btn_width)
        btn_export.grid(row=0, column=2, sticky="ew", padx=(17, 5))
        btn_export.tooltip = self.create_tooltip(btn_export, "Exporter les résultats filtrés vers Excel")

        data_frame = ttk.Frame(main_frame, style="Modern.TFrame")
        data_frame.grid(row=2, column=0, sticky="nsew")
        data_frame.columnconfigure(0, weight=1)
        data_frame.rowconfigure(0, weight=1)

        left_frame = ttk.Frame(data_frame, style="Modern.TFrame")
        left_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        left_frame.columnconfigure(0, weight=1)
        left_frame.rowconfigure(0, weight=1)

        self.create_treeview(left_frame)

        pagination_frame = ttk.Frame(left_frame, style="Modern.TFrame")
        pagination_frame.grid(row=1, column=0, sticky="ew", pady=(5, 0))
        self.create_pagination(pagination_frame)

        screen_width = self.root.winfo_screenwidth()
        right_panel_width = min(max(int(screen_width * 0.22), 320), 380)
        right_frame = ttk.Frame(data_frame, width=right_panel_width, style="Modern.TFrame")
        right_frame.grid(row=0, column=1, sticky="ns")
        right_frame.grid_propagate(False)

        details_frame = ttk.LabelFrame(right_frame, text="📋 Détails de la pièce", padding="15", style="Modern.TLabelframe")
        details_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        self.create_details_form(details_frame)

        image_frame = ttk.LabelFrame(right_frame, text="🖼️ Gestion des images", padding="15", style="Modern.TLabelframe")
        image_frame.pack(fill=tk.X, pady=(0, 10))
        self.create_image_section(image_frame)

        buttons_frame = ttk.Frame(right_frame, style="Modern.TFrame")
        buttons_frame.pack(fill=tk.X)
        self.create_action_buttons(buttons_frame)
        self.create_status_bar(main_frame)
        # Focus automatique sur le champ principal
        self.root.after(200, lambda: self.search_article.focus_set())

    def create_treeview(self, parent):
        tree_frame = ttk.Frame(parent)
        tree_frame.grid(row=0, column=0, sticky="nsew")
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)
        columns = ("ID", "Article", "Code SAP", "Description", "Description longue", "Unité", "Statut", "Quantité installée", "Situation", "Image ?")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=15, style="Modern.Treeview")
        tree_frame.columnconfigure(tuple(range(len(columns))), weight=1)
        tree_frame.rowconfigure(0, weight=1)
        screen_width = self.root.winfo_screenwidth()
        base_width = max(screen_width - 450, 700)
        column_widths = {
            "ID": 45,
            "Article": int(base_width * 0.09),
            "Code SAP": int(base_width * 0.09),
            "Description": int(base_width * 0.15),
            "Description longue": int(base_width * 0.18),
            "Unité": int(base_width * 0.07),
            "Statut": int(base_width * 0.07),
            "Quantité installée": int(base_width * 0.09),
            "Situation": int(base_width * 0.12),
            "Image ?": 70
        }
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=column_widths.get(col, 100), minwidth=60, stretch=True)
        v_scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        h_scrollbar = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        h_scrollbar.grid(row=1, column=0, sticky="ew")
        self.tree.bind("<<TreeviewSelect>>", self.on_item_select)
        self.tree.bind("<Double-1>", self.show_details_window)

    def create_pagination(self, parent):
        ttk.Button(parent, text="<<", command=self.first_page).pack(side=tk.LEFT, padx=2)
        ttk.Button(parent, text="<", command=self.prev_page).pack(side=tk.LEFT, padx=2)
        self.page_label = ttk.Label(parent, text="Page 1 / 1")
        self.page_label.pack(side=tk.LEFT, padx=10)
        ttk.Button(parent, text=">", command=self.next_page).pack(side=tk.LEFT, padx=2)
        ttk.Button(parent, text=">>", command=self.last_page).pack(side=tk.LEFT, padx=2)
        ttk.Label(parent, text="Taille:").pack(side=tk.LEFT, padx=(20, 5))
        self.page_size_var = tk.StringVar(value="100")
        page_size_combo = ttk.Combobox(parent, textvariable=self.page_size_var, values=["50", "100", "200", "500"], width=8)
        page_size_combo.pack(side=tk.LEFT, padx=2)
        page_size_combo.bind("<<ComboboxSelected>>", self.change_page_size)

        screen_height = self.root.winfo_screenheight()
        if screen_height <= 768:
            self.page_size = 50
            self.page_size_var.set("50")
        elif screen_height >= 1440:
            self.page_size = 200
            self.page_size_var.set("200")

    def create_details_form(self, parent):
        canvas = tk.Canvas(parent, highlightthickness=0, bg='white')
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        form_frame = scrollable_frame
        self.detail_vars = {}
        
        screen_width = self.root.winfo_screenwidth()
        entry_width = min(max(int(screen_width * 0.02), 25), 35)

        main_fields = [
            ("Article *", "article", "text", True),
            ("Code SAP", "code_sap", "text", False),
            ("Description", "description", "text", False),
            ("Unité de mesure", "unite", "text", False),
            ("Statut", "statut", "text", False),
            ("Quantité installée", "quantite_installee", "text", False),
            ("Situation", "situation", "text", False)
        ]
        
        row_idx = 0
        for label, var_name, field_type, required in main_fields:
            label_color = "#dc2626" if required else "#374151"
            ttk.Label(form_frame, text=label, foreground=label_color, font=("Segoe UI", 9, "bold")).grid(
                row=row_idx, column=0, sticky=tk.W, pady=(5, 2), padx=(5, 0))
            
            if field_type == "combo":
                var = tk.StringVar()
                widget = ttk.Combobox(form_frame, textvariable=var, values=["Actif", "Désactivé", "En attente", "Obsolète"],
                                      state="readonly", width=entry_width, style="Modern.TCombobox")
            else:
                var = tk.StringVar()
                widget = ttk.Entry(form_frame, textvariable=var, width=entry_width, style="Modern.TEntry")
            
            widget.grid(row=row_idx+1, column=0, sticky=tk.EW, pady=(0, 8), padx=(5, 5))
            self.detail_vars[var_name] = var
            row_idx += 2
        
        ttk.Separator(form_frame, orient='horizontal').grid(row=row_idx, column=0, sticky="ew", pady=(10, 15), padx=5)
        row_idx += 1
        
        ttk.Label(form_frame, text="Description longue", font=("Segoe UI", 9, "bold")).grid(
            row=row_idx, column=0, sticky=tk.W, pady=(0, 5), padx=(5, 0))
        row_idx += 1
        
        text_frame = ttk.Frame(form_frame)
        text_frame.grid(row=row_idx, column=0, sticky="ew", padx=(5, 5), pady=(0, 10))
        
        self.description_longue_text = tk.Text(text_frame, height=4, width=entry_width, wrap=tk.WORD, font=("Segoe UI", 9))
        text_scroll = ttk.Scrollbar(text_frame, orient="vertical", command=self.description_longue_text.yview)
        self.description_longue_text.configure(yscrollcommand=text_scroll.set)
        
        self.description_longue_text.pack(side="left", fill="both", expand=True)
        text_scroll.pack(side="right", fill="y")
        
        form_frame.columnconfigure(0, weight=1)
        
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

    def create_action_buttons(self, parent):
        main_buttons_frame = ttk.Frame(parent)
        main_buttons_frame.pack(fill=tk.X, pady=(0, 3))
        
        new_btn = ttk.Button(main_buttons_frame, text="+ Nouveau", command=self.new_record, style="Primary.TButton")
        new_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 1))
        
        edit_btn = ttk.Button(main_buttons_frame, text="✏ Modifier", command=self.edit_record, style="Action.TButton")
        edit_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=1)
        
        save_buttons_frame = ttk.Frame(parent)
        save_buttons_frame.pack(fill=tk.X, pady=(0, 3))
        
        save_btn = ttk.Button(save_buttons_frame, text="💾 Sauvegarder", command=self.save_changes, style="Success.TButton")
        save_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 1))
        
        cancel_btn = ttk.Button(save_buttons_frame, text="❌ Annuler", command=self.cancel_changes, style="Action.TButton")
        cancel_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=1)
        
        delete_frame = ttk.Frame(parent)
        delete_frame.pack(fill=tk.X, pady=(3,0))
        
        delete_btn = ttk.Button(delete_frame, text="🗑 Supprimer", command=self.delete_record, style="Danger.TButton")
        delete_btn.pack(fill=tk.X)
        
        for btn in [new_btn, edit_btn, save_btn, cancel_btn, delete_btn]:
            self.add_hover_effect(btn)
        
        self.action_buttons = {'new': new_btn, 'edit': edit_btn, 'save': save_btn, 'cancel': cancel_btn, 'delete': delete_btn}
        
    def add_hover_effect(self, widget):
        def on_enter(e): widget.configure(cursor="hand2")
        def on_leave(e): widget.configure(cursor="")
        widget.bind("<Enter>", on_enter)
        widget.bind("<Leave>", on_leave)

    def load_data(self):
        try:
            self.update_status("Chargement en cours...", "loading")
            self.progress_bar.start()
            
            filters = self.get_current_filters()
            results, total_count = self.db_manager.search_pieces(
                filters=filters, limit=self.page_size, offset=self.current_page * self.page_size
            )
            
            self.total_records = total_count
            self.update_treeview(results)
            self.update_pagination()
            
            if results:
                self.update_status(f"Chargement terminé", "success")
            else:
                self.update_status("Aucun résultat trouvé", "warning")
                
        except Exception as e:
            self.update_status(f"Erreur: {str(e)}", "error")
            messagebox.showerror("Erreur", f"Erreur lors du chargement: {str(e)}")
        finally:
            self.progress_bar.stop()

    def get_current_filters(self):
        filters = {}
        if self.search_article.get().strip(): filters['article'] = self.search_article.get().strip()
        # Recherche : si l'utilisateur tape 'vide' dans Code SAP, filtrer les codes SAP vides/None/NaN
        code_sap_val = self.search_sap.get().strip()
        if code_sap_val:
            if code_sap_val.lower() == 'vide':
                filters['code_sap_empty'] = True
            else:
                filters['code_sap'] = code_sap_val
        if self.search_description.get().strip(): filters['description'] = self.search_description.get().strip()
        if self.search_description_longue.get().strip(): filters['description_longue'] = self.search_description_longue.get().strip()
        if self.search_statut.get() != "Tous": filters['statut'] = self.search_statut.get()
        if self.search_unite.get() != "Tous": filters['unite'] = self.search_unite.get()
        if self.search_quantite_installee.get().strip(): filters['quantite_installee'] = self.search_quantite_installee.get().strip()
        if self.search_situation.get().strip(): filters['situation'] = self.search_situation.get().strip()
        return filters

    def update_treeview(self, results):
        for item in self.tree.get_children(): self.tree.delete(item)
        for idx, row in enumerate(results):
            # Affichage : remplacer NaN, 'nan', None ou '' par '' pour le code SAP
            row = list(row[:10])
            code_sap_val = row[2]
            if code_sap_val is None or (isinstance(code_sap_val, float) and math.isnan(code_sap_val)) or str(code_sap_val).lower() == 'nan':
                row[2] = ''
            # Ajout colonne Image ?
            image_path = row[9] if len(row) > 9 else None
            has_image = (image_path and os.path.exists(image_path))
            image_status = "✅" if has_image else "❌"
            display_row = row + [image_status]
            tag = 'oddrow' if idx % 2 else 'evenrow'
            self.tree.insert("", tk.END, values=display_row, tags=(tag,))
        self.tree.tag_configure('oddrow', background=self.treeview_row_colors[1])
        self.tree.tag_configure('evenrow', background=self.treeview_row_colors[0])

    def update_pagination(self):
        total_pages = max(1, (self.total_records + self.page_size - 1) // self.page_size)
        self.page_label.config(text=f"Page {self.current_page + 1} / {total_pages}")

    def search_data(self):
        self.current_page = 0
        self.load_data()

    def reset_search(self):
        self.search_article.delete(0, tk.END)
        self.search_sap.delete(0, tk.END)
        self.search_description.delete(0, tk.END)
        self.search_description_longue.delete(0, tk.END)
        self.search_statut.set("Tous")
        self.search_unite.set("Tous")
        self.search_quantite_installee.delete(0, tk.END)
        self.search_situation.delete(0, tk.END)
        self.current_page = 0
        self.load_data()

    def first_page(self): self.current_page = 0; self.load_data()
    def prev_page(self):
        if self.current_page > 0: self.current_page -= 1; self.load_data()
    def next_page(self):
        total_pages = (self.total_records + self.page_size - 1) // self.page_size
        if self.current_page < total_pages - 1: self.current_page += 1; self.load_data()
    def last_page(self):
        self.current_page = max(0, (self.total_records + self.page_size - 1) // self.page_size - 1); self.load_data()
    def change_page_size(self, event): self.page_size = int(self.page_size_var.get()); self.current_page = 0; self.load_data()

    def on_item_select(self, event):
        selection = self.tree.selection()
        if selection:
            values = self.tree.item(selection[0])["values"]
            if values:
                self.current_piece_id = values[0]
                self.load_piece_details_from_id(self.current_piece_id)
        else:
            self.current_piece_id = None
        self.update_button_states()
        self.update_status(self.status_bar.cget("text").split(" ", 1)[1])

    def load_piece_details_from_id(self, piece_id):
        if piece_id:
            piece_data = self.db_manager.get_piece_by_id(piece_id)
            if piece_data: self.load_piece_details(piece_data)
    
    def resize_image(self, image_path, max_size=(800, 600)):
        try:
            with Image.open(image_path) as img:
                if img.width > max_size[0] or img.height > max_size[1]:
                    img.thumbnail(max_size, Image.Resampling.LANCZOS)
                    img.save(image_path, optimize=True, quality=85)
        except Exception as e: print(f"Erreur lors du redimensionnement: {e}")

    def show_details_window(self, event):
        selection = self.tree.selection()
        if not selection: return
        item_id = self.tree.item(selection[0])["values"][0]
        piece_data = self.db_manager.get_piece_by_id(item_id)
        if not piece_data: messagebox.showerror("Erreur", "Impossible de récupérer les détails."); return
        details_win = tk.Toplevel(self.root); details_win.title(f"Détails : {piece_data[1]}")
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        detail_width = min(max(int(screen_width * 0.4), 500), 800)
        detail_height = min(max(int(screen_height * 0.7), 600), 900)
        details_win.geometry(f"{detail_width}x{detail_height}")
        details_win.minsize(500, 600)
        details_win.transient(self.root); details_win.grab_set()
        details_win.resizable(True, True)
        main_frame = ttk.Frame(details_win, padding="15"); main_frame.pack(fill=tk.BOTH, expand=True)
        # --- Champs principaux ---
        details_frame = ttk.LabelFrame(main_frame, text="Informations sur la pièce", padding="10"); details_frame.pack(fill=tk.X, pady=(0, 10), side=tk.TOP); details_frame.columnconfigure(1, weight=1)
        fields = {"ID": piece_data[0], "Article": piece_data[1], "Code SAP": piece_data[2], "Description": piece_data[3], "Unité de mesure": piece_data[5], "Statut": piece_data[6], "Quantité installée": piece_data[7], "Situation": piece_data[8]}
        for i, (label, value) in enumerate(fields.items()):
            ttk.Label(details_frame, text=f"{label} :", font=("Arial", 11, "bold")).grid(row=i, column=0, sticky="nw", pady=4, padx=5)
            ttk.Label(details_frame, text=value or "N/A", wraplength=400, anchor="w", font=("Arial", 11)).grid(row=i, column=1, sticky="ew", padx=10)
        # --- Description longue ---
        desc_longue_frame = ttk.LabelFrame(main_frame, text="Description longue", padding="10")
        desc_longue_frame.pack(fill=tk.X, pady=(0, 10), side=tk.TOP)
        desc_longue_text = tk.Text(desc_longue_frame, height=6, wrap=tk.WORD, font=("Arial", 11), relief=tk.SOLID, borderwidth=1)
        desc_longue_text.pack(fill=tk.BOTH, expand=True)
        desc_longue_text.insert(tk.END, piece_data[4] or "Non spécifiée")
        desc_longue_text.config(state="disabled")
        # --- Image associée ---
        image_frame = ttk.LabelFrame(main_frame, text="Image associée", padding="10")
        image_frame.pack(fill=tk.BOTH, expand=True, side=tk.TOP)
        image_label = ttk.Label(image_frame, text="Chargement...", anchor=tk.CENTER)
        image_label.pack(fill=tk.BOTH, expand=True)
        def load_image_in_thread(path, target_label):
            if path and os.path.exists(path):
                try:
                    image = Image.open(path)
                    image.thumbnail((550, 450), Image.Resampling.LANCZOS)
                    photo = ImageTk.PhotoImage(image)
                    target_label.config(image=photo, text="")
                    target_label.image = photo
                except:
                    target_label.config(image="", text="Erreur d'image")
            else:
                target_label.config(image="", text="Aucune image associée")
        threading.Thread(target=load_image_in_thread, args=(piece_data[9], image_label), daemon=True).start()
        # Focus automatique sur la fenêtre
        details_win.after(200, lambda: details_win.focus_force())

    def edit_record(self):
        if self.current_piece_id is None: messagebox.showwarning("Attention", "Veuillez sélectionner une pièce."); return
        piece_data = self.db_manager.get_piece_by_id(self.current_piece_id)
        if piece_data:
            self.load_piece_details(piece_data); self.editing_mode = True
            self.update_button_states()
            self.update_status("Mode édition", "info")
        else: messagebox.showerror("Erreur", f"Impossible de charger la pièce ID {self.current_piece_id}.")

    def export_to_excel(self):
        file_path = filedialog.asksaveasfilename(title="Exporter vers Excel", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            try:
                self.update_status("Exportation...", "loading"); self.progress_bar.start()
                count = self.db_manager.export_to_excel(file_path, self.get_current_filters())
                self.update_status("Export terminé.", "success")
                messagebox.showinfo("Export terminé", f"{count} enregistrements exportés vers:\n{file_path}")
            except Exception as e: messagebox.showerror("Erreur", f"Erreur d'export: {str(e)}"); self.update_status("Erreur export", "error")
            finally: self.progress_bar.stop()

    def create_status_bar(self, parent):
        status_main_frame = ttk.Frame(parent, style="Modern.TFrame")
        status_main_frame.grid(row=3, column=0, sticky="ew", pady=(14, 0))
        status_frame = ttk.Frame(status_main_frame, relief=tk.SUNKEN, borderwidth=1, style="Modern.TFrame")
        status_frame.grid(row=0, column=0, sticky="ew", pady=(0, 7))
        status_frame.columnconfigure(1, weight=1)
        self.status_bar = ttk.Label(status_frame, text="✅ Prêt", font=("Segoe UI", 10, "bold"), foreground="#10b981", background="#f8fafc", relief=tk.FLAT)
        self.status_bar.grid(row=0, column=0, padx=(14, 0), sticky="w")
        ttk.Separator(status_frame, orient='vertical').grid(row=0, column=1, sticky="ns", padx=10)
        self.info_label = ttk.Label(status_frame, text="", font=("Segoe UI", 10), foreground="#6b7280", background="#f8fafc")
        self.info_label.grid(row=0, column=2, padx=(0, 14), sticky="w")
        self.time_label = ttk.Label(status_frame, text="", font=("Segoe UI", 10), foreground="#6b7280", background="#f8fafc")
        self.time_label.grid(row=0, column=3, padx=(0, 14), sticky="e")
        self.update_time()
        self.progress_bar = ttk.Progressbar(status_main_frame, mode='indeterminate', style="Modern.Horizontal.TProgressbar")
        self.progress_bar.grid(row=1, column=0, sticky="ew", pady=(0, 7))

    def update_time(self):
        self.time_label.config(text=datetime.now().strftime("%H:%M:%S - %d/%m/%Y"))
        self.root.after(1000, self.update_time)

    def update_status(self, message, status_type="info"):
        icons = {"info": "ℹ️", "success": "✅", "warning": "⚠️", "error": "❌", "loading": "🔄"}
        colors = {"info": "#2563eb", "success": "#10b981", "warning": "#f59e0b", "error": "#ef4444", "loading": "#6b7280"}
        icon, color = icons.get(status_type, icons["info"]), colors.get(status_type, colors["info"])
        self.status_bar.config(text=f"{icon} {message}", foreground=color)
        
        if hasattr(self, 'total_records'):
            selected = len(self.tree.selection())
            info_text = f"Total: {self.total_records}" + (f" | Sélectionnés: {selected}" if selected > 0 else "")
            self.info_label.config(text=info_text)
        self.root.update_idletasks()
        
    def update_button_states(self):
        if not hasattr(self, 'action_buttons'): return
        is_item_selected = bool(self.current_piece_id)
        if self.editing_mode:
            for name, state in {'new': 'disabled', 'edit': 'disabled', 'save': 'normal', 'cancel': 'normal', 'delete': 'disabled'}.items():
                self.action_buttons[name].configure(state=state)
        else:
            self.action_buttons['new'].configure(state='normal')
            self.action_buttons['edit'].configure(state='normal' if is_item_selected else 'disabled')
            self.action_buttons['save'].configure(state='disabled')
            self.action_buttons['cancel'].configure(state='disabled')
            self.action_buttons['delete'].configure(state='normal' if is_item_selected else 'disabled')

    def setup_keyboard_shortcuts(self):
        self.root.bind('<Control-n>', lambda e: self.new_record())
        self.root.bind('<Control-e>', lambda e: self.edit_record())
        self.root.bind('<Control-s>', lambda e: self.save_changes())
        self.root.bind('<Control-d>', lambda e: self.delete_record())
        self.root.bind('<Escape>', lambda e: self.cancel_changes())
        self.root.bind('<Control-f>', lambda e: self.search_article.focus_set())
        self.root.bind('<F5>', lambda e: self.load_data())
        self.root.bind('<Control-Left>', lambda e: self.prev_page())
        self.root.bind('<Control-Right>', lambda e: self.next_page())
        self.root.bind('<Control-Home>', lambda e: self.first_page())
        self.root.bind('<Control-End>', lambda e: self.last_page())
        self.root.bind('<Control-Alt-e>', lambda e: self.export_to_excel())
        self.root.bind('<Control-i>', lambda e: self.load_image())
        self.root.bind('<Control-Delete>', lambda e: self.remove_image())
        for widget in [self.search_article, self.search_sap, self.search_description, self.search_quantite_installee, self.search_situation]:
            widget.bind('<Return>', lambda e: self.search_data())

    def create_help_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        # Ajout du menu Historique
        history_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Historique", menu=history_menu)
        history_menu.add_command(label="Afficher l'historique", command=self.show_history_window)
        # Menu Aide
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Aide", menu=help_menu)
        help_menu.add_command(label="Raccourcis clavier", command=self.show_shortcuts_window)
        help_menu.add_separator()
        help_menu.add_command(label="À propos", command=self.show_about)

    def show_shortcuts_window(self):
        win = tk.Toplevel(self.root); win.title("Raccourcis clavier")
        
        screen_width = self.root.winfo_screenwidth()
        shortcut_width = min(max(int(screen_width * 0.3), 400), 600)
        win.geometry(f"{shortcut_width}x450")
        
        win.resizable(False, False); win.transient(self.root); win.grab_set()
        frame = ttk.Frame(win, padding="20"); frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frame, text="Raccourcis clavier", font=("Segoe UI", 16, "bold")).pack(pady=(0, 20))
        tree = ttk.Treeview(frame, columns=("Raccourci", "Action"), show="headings", height=15)
        tree.heading("Raccourci", text="Raccourci"); tree.heading("Action", text="Action"); tree.column("Raccourci", width=120); tree.column("Action", width=280)
        shortcuts = [
            ("Ctrl+N", "Nouveau"), ("Ctrl+E", "Modifier"), ("Ctrl+S", "Sauvegarder"), ("Ctrl+D / Ctrl+Suppr", "Supprimer pièce/image"),
            ("Échap", "Annuler"), ("Ctrl+F", "Focus Recherche"), ("F5", "Actualiser les données"), ("Ctrl+Gauche/Droite", "Page préc./suiv."),
            ("Ctrl+Début/Fin", "Première/Dernière page"), ("Ctrl+Alt+E", "Exporter Excel"), ("Ctrl+I", "Charger image"), ("Entrée", "Rechercher")
        ]
        for sc, act in shortcuts: tree.insert("", tk.END, values=(sc, act))
        tree.pack(fill=tk.BOTH, expand=True)
        ttk.Button(frame, text="Fermer", command=win.destroy).pack(pady=(20, 0))

    def show_about(self):
        messagebox.showinfo("À propos", "Gestionnaire de Pièces OCP \n\n© 2025 - By Ayman Djioui")

    def show_history_window(self):
        win = tk.Toplevel(self.root)
        win.title("Historique des modifications")
        screen_width = self.root.winfo_screenwidth()
        hist_width = min(max(int(screen_width * 0.35), 500), 800)
        win.geometry(f"{hist_width}x500")
        win.resizable(True, True)
        win.transient(self.root)
        win.grab_set()
        frame = ttk.Frame(win, padding="15")
        frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frame, text="Historique des modifications", font=("Segoe UI", 16, "bold")).pack(pady=(0, 10))
        text = tk.Text(frame, wrap=tk.WORD, height=25, state="normal", font=("Segoe UI", 10))
        text.pack(fill=tk.BOTH, expand=True)
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        text.config(yscrollcommand=scrollbar.set)
        try:
            with open(self.HISTORIQUE_FILE, 'r', encoding='utf-8') as f:
                content = f.read()
        except Exception as e:
            content = f"Erreur de lecture de l'historique : {e}"
        text.insert(tk.END, content)
        text.config(state="disabled")
        ttk.Button(frame, text="Fermer", command=win.destroy).pack(pady=(10, 0))

    def on_closing(self):
        if self.editing_mode and messagebox.askyesno("Confirmation", "Des modifications sont en cours. Fermer?"):
            self.executor.shutdown(wait=False); self.root.destroy()
        elif not self.editing_mode:
            self.executor.shutdown(wait=False); self.root.destroy()

    def create_image_section(self, parent):
        image_buttons = ttk.Frame(parent)
        image_buttons.pack(fill=tk.X, pady=5)
        
        load_btn = ttk.Button(image_buttons, text="📁 Ajouter", command=self.load_image, style="Primary.TButton")
        load_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 1))
        
        remove_btn = ttk.Button(image_buttons, text="🗑 Supprimer", command=self.remove_image, style="Danger.TButton")
        remove_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=1)
        
        self.add_hover_effect(load_btn)
        self.add_hover_effect(remove_btn)

    def load_piece_details(self, piece_data):
        def safe_str(val):
            if val is None:
                return ""
            if isinstance(val, float) and math.isnan(val):
                return ""
            if str(val).lower() == "nan":
                return ""
            return str(val)
        self.detail_vars["article"].set(safe_str(piece_data[1]))
        self.detail_vars["code_sap"].set(safe_str(piece_data[2]))
        self.detail_vars["description"].set(safe_str(piece_data[3]))
        self.detail_vars["unite"].set(safe_str(piece_data[5]))
        self.detail_vars["statut"].set(safe_str(piece_data[6]))
        self.detail_vars["quantite_installee"].set(safe_str(piece_data[7]))
        self.detail_vars["situation"].set(safe_str(piece_data[8]))
        self.description_longue_text.delete(1.0, tk.END)
        self.description_longue_text.insert(1.0, safe_str(piece_data[4]))
        self.current_image = piece_data[9]

    def new_record(self):
        self.editing_mode = True
        self.current_piece_id = None
        for var in self.detail_vars.values():
            var.set("")
        self.description_longue_text.delete(1.0, tk.END)
        self.current_image = None
        self.detail_vars["statut"].set("Actif")
        self.detail_vars["quantite_installee"].set("")
        self.detail_vars["situation"].set("")
        self.update_button_states()
        self.update_status("Mode création", "info")
        self.log_history("Création", self.current_piece_id, f"Article: {self.detail_vars['article'].get()}")

    def delete_record(self):
        if self.current_piece_id is None:
            messagebox.showwarning("Attention", "Veuillez sélectionner une pièce.")
            return
        
        if messagebox.askyesno("Confirmation", f"Supprimer la pièce ID {self.current_piece_id}?\nCette action est irréversible."):
            try:
                piece_data = self.db_manager.get_piece_by_id(self.current_piece_id)
                if piece_data and piece_data[9] and os.path.exists(piece_data[9]):
                    os.remove(piece_data[9])
                
                self.db_manager.delete_piece(self.current_piece_id)
                # Historique détaillé suppression
                champs = ["Article", "Code SAP", "Description", "Description longue", "Unité de mesure", "Statut", "Quantité installée", "Situation", "Image"]
                old_data = {champs[i]: piece_data[i+1] for i in range(len(champs))} if piece_data else None
                self.log_history("Suppression", self.current_piece_id, f"Article: {piece_data[1] if piece_data else ''}", old_data=old_data, new_data=None)
                self.current_piece_id = None
                
                for var in self.detail_vars.values():
                    var.set("")
                self.description_longue_text.delete(1.0, tk.END)
                self.current_image = None
                
                self.update_status("Pièce supprimée.", "success")
                self.load_data()
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur de suppression: {str(e)}")
                self.update_status("Erreur suppression", "error")

    def cancel_changes(self):
        if self.editing_mode and messagebox.askyesno("Confirmation", "Voulez-vous annuler les modifications?"):
            self.editing_mode = False
            self.update_button_states()
            
            if self.current_piece_id:
                piece_data = self.db_manager.get_piece_by_id(self.current_piece_id)
                if piece_data:
                    self.load_piece_details(piece_data)
            else:
                for var in self.detail_vars.values():
                    var.set("")
                self.description_longue_text.delete(1.0, tk.END)
                self.current_image = None
                self.detail_vars["quantite_installee"].set("")
                self.detail_vars["situation"].set("")
            
            self.update_status("Modifications annulées", "warning")
            
    def remove_image(self):
        if self.current_piece_id is None:
            messagebox.showwarning("Attention", "Veuillez sélectionner une pièce d'abord")
            return
        
        if messagebox.askyesno("Confirmation", "Voulez-vous vraiment supprimer l'image de cette pièce?"):
            try:
                self.current_image = None
                self.log_history("Suppression image", self.current_piece_id)
                
                if not self.editing_mode:
                    self.editing_mode = True
                    self.update_button_states()
                    self.update_status("Mode édition activé - Image supprimée", "info")
                else:
                    self.update_status("Image supprimée", "success")
                    
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur de suppression: {str(e)}")
                self.update_status("Erreur image", "error")
                
    def load_image(self, *args):
        if self.current_piece_id is None:
            messagebox.showwarning("Attention", "Veuillez sélectionner une pièce d'abord")
            return
        
        file_path = filedialog.askopenfilename(
            title="Sélectionner une image",
            filetypes=[("Images", "*.jpg *.jpeg *.png"), ("Tous les fichiers", "*.*")]
        )
        
        if file_path:
            try:
                self.current_image = file_path
                self.log_history("Ajout image", self.current_piece_id, f"Fichier: {os.path.basename(file_path)}")
                
                if not self.editing_mode:
                    self.editing_mode = True
                    self.update_button_states()
                    self.update_status("Mode édition activé - Image prête à être ajoutée", "info")
                else:
                    self.update_status("Image prête à être ajoutée", "success")
                    
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur lors de la sélection de l'image: {str(e)}")
                self.update_status("Erreur image", "error")

    def save_changes(self):
        if not self.editing_mode:
            messagebox.showwarning("Attention", "Aucune modification en cours")
            return
        if not self.detail_vars["article"].get().strip():
            messagebox.showerror("Erreur", "Le champ Article est obligatoire")
            return
        try:
            final_image_path = self.current_image
            old_piece_data = None
            if self.current_piece_id:
                 old_piece_data = self.db_manager.get_piece_by_id(self.current_piece_id)
            if self.current_image and not self.current_image.startswith(self.images_folder):
                if os.path.exists(self.current_image):
                    new_filename = f"piece_{self.current_piece_id or 'new'}_{datetime.now().strftime('%Y%m%d_%H%M%S')}{os.path.splitext(self.current_image)[1]}"
                    final_image_path = os.path.join(self.images_folder, new_filename)
                    shutil.copy2(self.current_image, final_image_path)
                    self.resize_image(final_image_path)
                else:
                    messagebox.showwarning("Image manquante", "L'image sélectionnée n'existe plus. Seule la fiche sera sauvegardée.")
                    final_image_path = ""
            if old_piece_data and old_piece_data[9] and old_piece_data[9] != final_image_path:
                if os.path.exists(old_piece_data[9]):
                    try: os.remove(old_piece_data[9])
                    except: pass
            def clean_code_sap(val):
                if val is None:
                    return ""
                if isinstance(val, float) and math.isnan(val):
                    return ""
                if str(val).lower() == "nan":
                    return ""
                return str(val)
            piece_data = (
                self.detail_vars["article"].get().strip(),
                clean_code_sap(self.detail_vars["code_sap"].get().strip()),
                self.detail_vars["description"].get().strip(),
                self.description_longue_text.get(1.0, tk.END).strip(),
                self.detail_vars["unite"].get().strip(),
                self.detail_vars["statut"].get(),
                self.detail_vars["quantite_installee"].get().strip(),
                self.detail_vars["situation"].get().strip(),
                final_image_path or ""
            )
            champs = ["Article", "Code SAP", "Description", "Description longue", "Unité de mesure", "Statut", "Quantité installée", "Situation", "Image"]
            if self.current_piece_id is None:
                new_id = self.db_manager.insert_piece(piece_data)
                self.current_piece_id = new_id
                if final_image_path and 'new' in final_image_path:
                    new_final_path = final_image_path.replace('new', str(new_id))
                    if os.path.exists(final_image_path):
                        os.rename(final_image_path, new_final_path)
                        self.db_manager.update_piece(new_id, piece_data[:-1] + (new_final_path,))
                self.update_status(f"Pièce {new_id} créée.", "success")
                # Historique détaillé création
                new_data = {champs[i]: piece_data[i] for i in range(len(champs))}
                self.log_history("Création", new_id, f"Article: {piece_data[0]}", old_data=None, new_data=new_data)
            else:
                # Historique détaillé modification
                old_data = {champs[i]: old_piece_data[i+1] for i in range(len(champs))} if old_piece_data else None
                new_data = {champs[i]: piece_data[i] for i in range(len(champs))}
                self.db_manager.update_piece(self.current_piece_id, piece_data)
                self.update_status(f"Pièce {self.current_piece_id} mise à jour.", "success")
                self.log_history("Modification", self.current_piece_id, f"Article: {piece_data[0]}", old_data=old_data, new_data=new_data)
            self.editing_mode = False
            self.update_button_states()
            self.load_data()
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur de sauvegarde: {str(e)}")
            self.update_status("Erreur sauvegarde", "error")

    def create_tooltip(self, widget, text):
        # Simple tooltip pour les boutons
        tooltip = tk.Toplevel(widget)
        tooltip.withdraw()
        tooltip.overrideredirect(True)
        label = tk.Label(tooltip, text=text, background="#f8fafc", relief="solid", borderwidth=1, font=("Segoe UI", 9))
        label.pack()
        def enter(event):
            x = widget.winfo_rootx() + 40
            y = widget.winfo_rooty() + 20
            tooltip.geometry(f"+{x}+{y}")
            tooltip.deiconify()
        def leave(event):
            tooltip.withdraw()
        widget.bind("<Enter>", enter)
        widget.bind("<Leave>", leave)
        return tooltip

def main():
    root = tk.Tk()
    root.iconbitmap(resource_path('ocp.ico')) 
    app = OCPPiecesManager(root)  
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
    y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
    root.geometry(f'+{x}+{y}')
    root.mainloop()

if __name__ == "__main__":
    main()