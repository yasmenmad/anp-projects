import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from datetime import datetime
import pytz
from difflib import get_close_matches
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import threading
from tkinter import font as tkfont

class EnhancedBadgeApp:
    def __init__(self, root):  # Changé de _init à _init_
        self.root = root
        self.root.title("ANP - Système Intelligent de Gestion des Badges")
        self.root.geometry("1200x800")
        self.root.configure(bg="#f0f2f5")
        
# Correction de la ligne 14 (méthode _init_)



        self.setup_styles()
        self.load_data()
        self.create_widgets()
        self.refresh()
        self.current_selection = None
    
    def setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TFrame', background="#f0f2f5")
        style.configure('TLabel', background="#f0f2f5", font=('Helvetica', 10))
        style.configure('TButton', font=('Helvetica', 10), padding=6)
        style.configure('TEntry', font=('Helvetica', 10), padding=5)
        style.configure('Header.TLabel', font=('Helvetica', 12, 'bold'))
        style.configure('Treeview', font=('Helvetica', 9), rowheight=25)
        style.configure('Treeview.Heading', font=('Helvetica', 10, 'bold'))
        style.map('TButton', background=[('active', '#e1e5ea')])
        
        self.details_font = tkfont.Font(family="Courier New", size=10)
        self.title_font = tkfont.Font(family="Helvetica", size=12, weight="bold")
    
    def load_data(self):
        try:
            self.df = pd.read_excel("BADGES.xlsx")
            self.df.columns = [col.strip().replace(' ', '_') for col in self.df.columns]
            
            date_columns = ['ID_Modify_Time', 'Load_Date', 'Token_Modify_Time', 
                          'Issue_Date', 'Activation_Date', 'Deactivation_Date']
            
            for col in date_columns:
                if col in self.df.columns:
                    self.df[col] = pd.to_datetime(self.df[col], errors='coerce')
            
            self.df['Full_Name'] = self.df['First_Name'].str.strip() + ' ' + self.df['Last_Name'].str.strip()
                    
            messagebox.showinfo("Succès", f"Données chargées avec succès!\n{len(self.df)} enregistrements trouvés.")
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de charger le fichier:\n{e}")
            self.root.destroy()
    
    def create_widgets(self):
        header_frame = ttk.Frame(self.root, style='Header.TFrame')
        header_frame.pack(fill=tk.X, padx=10, pady=10)
        ttk.Label(header_frame, text="ANP - Système Intelligent de Gestion des Badges", 
                 style='Header.TLabel').pack(side=tk.LEFT)
        
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        search_frame = ttk.LabelFrame(main_frame, text="Recherche Intelligente", padding=15)
        search_frame.pack(fill=tk.X, pady=5)
        
        self.search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=50, font=('Helvetica', 11))
        search_entry.pack(side=tk.LEFT, padx=5)
        
        search_btn = ttk.Button(search_frame, text="Rechercher", command=self.search)
        search_btn.pack(side=tk.LEFT, padx=5)
        
        ai_btn = ttk.Button(search_frame, text="Assistant IA", command=self.open_ai_chat)
        ai_btn.pack(side=tk.RIGHT, padx=5)
        
        paned_window = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
        paned_window.pack(fill=tk.BOTH, expand=True)
        
        results_frame = ttk.Frame(paned_window)
        self.tree = ttk.Treeview(results_frame, columns=("ID", "Nom", "Prénom", "Numéro", "Statut", "Expiration"), 
                               show="headings", selectmode='browse')
        
        vsb = ttk.Scrollbar(results_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(results_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        
        results_frame.grid_rowconfigure(0, weight=1)
        results_frame.grid_columnconfigure(0, weight=1)
        
        columns = {
            "ID": {"width": 120, "anchor": tk.CENTER},
            "Nom": {"width": 150, "anchor": tk.W},
            "Prénom": {"width": 150, "anchor": tk.W},
            "Numéro": {"width": 100, "anchor": tk.CENTER},
            "Statut": {"width": 80, "anchor": tk.CENTER},
            "Expiration": {"width": 120, "anchor": tk.CENTER}
        }
        
        for col, params in columns.items():
            self.tree.heading(col, text=col)
            self.tree.column(col, **params)
        
        paned_window.add(results_frame)
        
        notebook = ttk.Notebook(paned_window)
        
        details_frame = ttk.Frame(notebook)
        self.details_text = scrolledtext.ScrolledText(details_frame, wrap=tk.WORD, 
                                                    font=self.details_font, state='disabled')
        self.details_text.pack(fill=tk.BOTH, expand=True)
        notebook.add(details_frame, text="Détails Complets")
        
        stats_frame = ttk.Frame(notebook)
        self.setup_stats_tab(stats_frame)
        notebook.add(stats_frame, text="Statistiques")
        
        paned_window.add(notebook)
        
        toolbar_frame = ttk.Frame(main_frame)
        toolbar_frame.pack(fill=tk.X, pady=5)
        
        buttons = [
            ("Actualiser", self.refresh),
            ("Graphiques", self.show_graphs),
            ("Quitter", self.root.quit)
        ]
        
        for text, command in buttons:
            btn = ttk.Button(toolbar_frame, text=text, command=command)
            btn.pack(side=tk.LEFT, padx=3)
        
        self.status_var = tk.StringVar()
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN)
        status_bar.pack(fill=tk.X)
        
        self.tree.bind("<Double-1>", self.show_details)
        self.tree.bind("<ButtonRelease-1>", self.show_details)
        search_entry.bind("<Return>", lambda e: self.search())
    
    def setup_stats_tab(self, frame):
        self.figure = plt.Figure(figsize=(5, 4), dpi=100)
        self.canvas = FigureCanvasTkAgg(self.figure, master=frame)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        self.update_stats()
    
    def update_stats(self):
        if not hasattr(self, 'df'):
            return
            
        self.figure.clear()
        
        ax1 = self.figure.add_subplot(121)
        status_counts = self.df['Token_Status'].value_counts()
        wedges, texts, autotexts = ax1.pie(status_counts, 
                                         colors=['#4CAF50', '#F44336'], 
                                         startangle=90,
                                         autopct=lambda p: f"{p:.1f}%\n({int(p/100*sum(status_counts))})")
        
        ax1.set_title('Répartition par statut')
        ax1.legend(['Actif', 'Inactif'], loc='upper right')
        
        ax2 = self.figure.add_subplot(122)
        if 'Deactivation_Date' in self.df.columns:
            today = datetime.now(pytz.utc)
            days_left = (self.df['Deactivation_Date'] - today).dt.days
            expiring_soon = days_left.between(0, 30).sum()
            not_expiring = len(self.df) - expiring_soon
            
            colors = ['#FFC107', '#2196F3']
            bars = ax2.bar(['Expirent bientôt', 'Autres'], 
                         [expiring_soon, not_expiring], 
                         color=colors)
            ax2.set_title('Badges expirant dans 30 jours')
            
            for bar in bars:
                height = bar.get_height()
                ax2.text(bar.get_x() + bar.get_width()/2., height,
                        f"{int(height)}", ha='center', va='bottom')
        
        self.figure.tight_layout()
        self.canvas.draw()
    
    def search(self):
        search_term = self.search_var.get().lower().strip()
        
        if not search_term:
            self.refresh()
            return
            
        mask = (self.df['Full_Name'].str.lower().str.contains(search_term, na=False)) | \
               (self.df['First_Name'].str.lower().str.contains(search_term, na=False)) | \
               (self.df['Last_Name'].str.lower().str.contains(search_term, na=False)) | \
               (self.df['External_System_ID'].astype(str).str.contains(search_term, na=False)) | \
               (self.df['Internal_Number'].astype(str).str.contains(search_term, na=False))
        
        results = self.df[mask]
        
        if len(results) == 0 and ' ' in search_term:
            parts = search_term.split()
            if len(parts) == 2:
                mask = ((self.df['First_Name'].str.lower().str.contains(parts[0], na=False)) & 
                       (self.df['Last_Name'].str.lower().str.contains(parts[1], na=False))) | \
                      ((self.df['First_Name'].str.lower().str.contains(parts[1], na=False)) & 
                       (self.df['Last_Name'].str.lower().str.contains(parts[0], na=False)))
                results = self.df[mask]
        
        if len(results) < 3 and len(self.df) > 10:
            suggestions = self.get_suggestions(search_term)
            if suggestions:
                self.status_var.set(f"Suggestions: {', '.join(suggestions[:3])}")
        
        self.display_results(results)
    
    def get_suggestions(self, term):
        all_names = self.df['First_Name'].str.lower().dropna().tolist() + \
                   self.df['Last_Name'].str.lower().dropna().tolist() + \
                   self.df['Full_Name'].str.lower().dropna().tolist()
        unique_names = list(set(all_names))
        return get_close_matches(term, unique_names, n=3, cutoff=0.6)
    
    def display_results(self, df):
        self.tree.delete(*self.tree.get_children())
        
        for _, row in df.iterrows():
            expiry = row.get('Deactivation_Date', '')
            if pd.notna(expiry):
                expiry = expiry.strftime('%d/%m/%Y')
            
            status = "Actif" if row.get('Token_Status', 0) == 1 else "Inactif"
            
            self.tree.insert("", tk.END, values=(
                row.get('External_System_ID', ''),
                row.get('Last_Name', ''),
                row.get('First_Name', ''),
                row.get('Internal_Number', ''),
                status,
                expiry
            ), iid=str(row.name))
    
    def show_details(self, event):
        item = self.tree.identify_row(event.y)
        if not item:
            return
            
        df_index = int(item)
        self.current_selection = df_index
        
        try:
            person_data = self.df.loc[df_index]
        except KeyError:
            return
            
        details = "\n"
        details += "╔══════════════════════════════════════════════════════════════════╗\n"
        details += "║            INFORMATIONS COMPLÈTES DU BADGE - ANP                ║\n"
        details += "╠══════════════════════════════════════════════════════════════════╣\n"
        details += "║                                                                 ║\n"
        
        details += "║  ■ IDENTITÉ                                                    ║\n"
        details += "║  ┌────────────────────────────────────────────────────────────┐ ║\n"
        details += f"║  │ {'Nom complet:':<15} {self.format_field(person_data.get('First_Name', '') + ' ' + person_data.get('Last_Name', ''), 40)} │ ║\n"
        details += f"║  │ {'ID:':<15} {self.format_field(person_data.get('External_System_ID', 'N/A'), 40)} │ ║\n"
        details += f"║  │ {'Numéro interne:':<15} {self.format_field(person_data.get('Internal_Number', 'N/A'), 40)} │ ║\n"
        details += f"║  │ {'Adresse:':<15} {self.format_field(person_data.get('Address', 'N/A'), 40)} │ ║\n"
        details += f"║  │ {'Rôles:':<15} {self.format_field(person_data.get('Roles', 'N/A'), 40)} │ ║\n"
        details += "║  └────────────────────────────────────────────────────────────┘ ║\n"
        details += "║                                                                 ║\n"
        
        details += "║  ■ STATUT                                                      ║\n"
        details += "║  ┌────────────────────────────────────────────────────────────┐ ║\n"
        details += f"║  │ {'Statut du badge:':<15} {self.format_field('Actif' if person_data.get('Token_Status', 0) == 1 else 'Inactif', 40)} │ ║\n"
        details += f"║  │ {'VIP:':<15} {self.format_field('Oui' if person_data.get('VIP', False) else 'Non', 40)} │ ║\n"
        details += f"║  │ {'Niveau d\'émission:':<15} {self.format_field(person_data.get('Issue_Level', 'N/A'), 40)} │ ║\n"
        details += "║  └────────────────────────────────────────────────────────────┘ ║\n"
        details += "║                                                                 ║\n"
        
        details += "║  ■ DATES IMPORTANTES                                           ║\n"
        details += "║  ┌────────────────────────────────────────────────────────────┐ ║\n"
        details += f"║  │ {'Émis le:':<15} {self.format_field(self.format_date(person_data.get('Issue_Date')), 40)} │ ║\n"
        details += f"║  │ {'Activé le:':<15} {self.format_field(self.format_date(person_data.get('Activation_Date')), 40)} │ ║\n"
        expiry = person_data.get('Deactivation_Date')
        details += f"║  │ {'Expire le:':<15} {self.format_field(self.format_date(expiry), 40)} │ ║\n"
        
        if pd.notna(expiry):
            today = datetime.now(pytz.utc)
            days_left = (expiry - today).days
            status = "⚠ Expire bientôt!" if days_left <= 30 else "✅ Valide"
            details += f"║  │ {'Jours restants:':<15} {self.format_field(f'{days_left} jours ({status})', 40)} │ ║\n"
        
        details += "║  └────────────────────────────────────────────────────────────┘ ║\n"
        details += "║                                                                 ║\n"
        details += "╚══════════════════════════════════════════════════════════════════╝\n"
        
        self.details_text.config(state='normal')
        self.details_text.delete(1.0, tk.END)
        self.details_text.insert(tk.END, details)
        self.details_text.config(state='disabled')
    
    def format_field(self, value, width):
        if value is None:
            value = 'N/A'
        return f"{str(value)[:width]:<{width}}"
    
    def format_date(self, date):
        if pd.isna(date):
            return "N/A"
        return date.strftime('%d/%m/%Y %H:%M')
    
    def refresh(self):
        if hasattr(self, 'df'):
            self.display_results(self.df)
            self.search_var.set("")
            self.details_text.config(state='normal')
            self.details_text.delete(1.0, tk.END)
            self.details_text.config(state='disabled')
            self.update_stats()
            self.status_var.set("Données actualisées")
            self.current_selection = None
        
    def show_graphs(self):
        try:
            if not hasattr(self, 'df') or self.df.empty:
                messagebox.showerror("Erreur", "Aucune donnée disponible pour les graphiques")
                return

            graph_window = tk.Toplevel(self.root)
            graph_window.title("Statistiques Avancées")
            graph_window.geometry("1100x750")
            
            # Vérifier les styles disponibles et utiliser un style de repli
            available_styles = plt.style.available
            preferred_styles = ['seaborn-v0_8', 'seaborn', 'ggplot', 'classic']
            selected_style = next((style for style in preferred_styles if style in available_styles), 'classic')
            plt.style.use(selected_style)
            
            plt.rcParams.update({
                'font.size': 10,
                'axes.titlesize': 12,
                'axes.labelsize': 10
            })

            notebook = ttk.Notebook(graph_window)
            notebook.pack(fill=tk.BOTH, expand=True)

            # Onglet 1 - Statut des badges
            tab1 = ttk.Frame(notebook)
            fig1 = plt.Figure(figsize=(6, 5), dpi=100)
            ax1 = fig1.add_subplot(111)
            
            status_counts = self.df['Token_Status'].value_counts()
            labels = ['Actif' if x == 1 else 'Inactif' for x in status_counts.index]
            colors = ['#4CAF50', '#F44336']  # Vert/Rouge
            
            wedges, _, _ = ax1.pie(
                status_counts,
                labels=labels,
                colors=colors,
                autopct=lambda p: f"{p:.1f}%\n({int(p/100*sum(status_counts))})",
                startangle=90,
                textprops={'fontsize': 10, 'color': 'white', 'weight': 'bold'},
                wedgeprops={'linewidth': 1, 'edgecolor': 'white'},
                pctdistance=0.85
            )
            
            ax1.set_title('Statut des badges', fontweight='bold', pad=20)
            ax1.legend(wedges, labels, title="Légende", loc="center left", bbox_to_anchor=(1, 0.5))
            
            canvas1 = FigureCanvasTkAgg(fig1, master=tab1)
            canvas1.draw()
            canvas1.get_tk_widget().pack(fill=tk.BOTH, expand=True)
            notebook.add(tab1, text="Statut")

            # Onglet 2 - Évolution temporelle
            tab2 = ttk.Frame(notebook)
            fig2 = plt.Figure(figsize=(8, 5), dpi=100)
            ax2 = fig2.add_subplot(111)
            
            self.df['Issue_Date'] = pd.to_datetime(self.df['Issue_Date'], errors='coerce', dayfirst=True)
            monthly = self.df.dropna(subset=['Issue_Date']).set_index('Issue_Date').resample('M').size()
            
            if not monthly.empty:
                monthly.plot(
                    ax=ax2,
                    marker='o',
                    color='#2196F3',
                    linestyle='-',
                    linewidth=2,
                    markersize=6
                )
                
                ax2.set_title('Badges émis par mois', fontweight='bold')
                ax2.set_xlabel('Mois', labelpad=10)
                ax2.set_ylabel('Nombre de badges', labelpad=10)
                ax2.grid(True, linestyle=':', alpha=0.6)
                fig2.autofmt_xdate(rotation=45)
            
            canvas2 = FigureCanvasTkAgg(fig2, master=tab2)
            canvas2.draw()
            canvas2.get_tk_widget().pack(fill=tk.BOTH, expand=True)
            notebook.add(tab2, text="Évolution")

            # Onglet 3 - Types de badges (si colonne existe)
            if 'Type' in self.df.columns:
                tab3 = ttk.Frame(notebook)
                fig3 = plt.Figure(figsize=(8, 6), dpi=100)
                ax3 = fig3.add_subplot(111)
                
                type_counts = self.df['Type'].value_counts()
                if not type_counts.empty:
                    colors = plt.cm.tab20(range(len(type_counts)))
                    bars = ax3.bar(
                        type_counts.index.astype(str),
                        type_counts.values,
                        color=colors,
                        edgecolor='grey',
                        linewidth=0.5
                    )
                    
                    ax3.set_title('Types de badges', fontweight='bold')
                    ax3.set_xlabel('Type de badge')
                    ax3.set_ylabel('Nombre')
                    ax3.grid(axis='y', linestyle=':', alpha=0.6)
                    
                    for bar in bars:
                        height = bar.get_height()
                        ax3.text(bar.get_x() + bar.get_width()/2., height,
                                f"{height}", ha='center', va='bottom', fontsize=9)
                
                canvas3 = FigureCanvasTkAgg(fig3, master=tab3)
                canvas3.draw()
                canvas3.get_tk_widget().pack(fill=tk.BOTH, expand=True)
                notebook.add(tab3, text="Types")

        except Exception as e:
            messagebox.showerror("Erreur Graphique", f"Erreur lors de la génération des graphiques:\n{str(e)}")
            if 'graph_window' in locals():
                graph_window.destroy()
    def open_ai_chat(self):
        self.chat_window = tk.Toplevel(self.root)
        self.chat_window.title("Assistant Virtuel ANP")
        self.chat_window.geometry("600x700")

        chat_frame = ttk.Frame(self.chat_window)
        chat_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        suggested_frame = ttk.LabelFrame(chat_frame, text="Questions suggérées", padding=10)
        suggested_frame.pack(fill=tk.X, pady=(0, 10))

        suggested_questions = [
            "Statut des badges",
            "Liste des badges expirés",
            "Badges expirant ce mois",
            "Nombre de VIP",
            "Statistiques complètes"
        ]

        for question in suggested_questions:
            btn = ttk.Button(suggested_frame, text=question,
                             command=lambda q=question: self.insert_question(q))
            btn.pack(side=tk.TOP, fill=tk.X, pady=2)

        self.chat_history = scrolledtext.ScrolledText(chat_frame, wrap=tk.WORD,
                                                      font=('Helvetica', 10), state='disabled')
        self.chat_history.pack(fill=tk.BOTH, expand=True)

        input_frame = ttk.Frame(chat_frame)
        input_frame.pack(fill=tk.X, pady=5)

        self.user_input = ttk.Entry(input_frame)
        self.user_input.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        send_btn = ttk.Button(input_frame, text="Envoyer", command=self.send_ai_message)
        send_btn.pack(side=tk.RIGHT)

        self.add_to_chat("Assistant: Bonjour! Je suis l'assistant ANP pour la gestion des badges.\n"
                         "Je peux vous fournir:\n"
                         "- Des statistiques sur les badges\n"
                         "- Des listes de badges spécifiques\n"
                         "- Des informations détaillées\n\n"
                         "Posez-moi votre question ou cliquez sur une suggestion ci-dessus.")

    def insert_question(self, question):
        self.user_input.delete(0, tk.END)
        self.user_input.insert(0, question)

    def add_to_chat(self, message):
        self.chat_history.config(state='normal')
        self.chat_history.insert(tk.END, message + "\n\n")
        self.chat_history.config(state='disabled')
        self.chat_history.see(tk.END)

    def send_ai_message(self):
        user_msg = self.user_input.get()
        if not user_msg:
            return

        self.add_to_chat(f"Vous: {user_msg}")
        self.user_input.delete(0, tk.END)

        threading.Thread(target=self.process_ai_response, args=(user_msg,), daemon=True).start()

    def process_ai_response(self, user_msg):
        response = self.generate_ai_response(user_msg)
        self.chat_window.after(0, self.add_to_chat, f"Assistant: {response}")

    def open_ai_chat(self):
        self.chat_window = tk.Toplevel(self.root)
        self.chat_window.title("Assistant Virtuel ANP")
        self.chat_window.geometry("600x700")

        chat_frame = ttk.Frame(self.chat_window)
        chat_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        suggested_frame = ttk.LabelFrame(chat_frame, text="Questions suggérées", padding=10)
        suggested_frame.pack(fill=tk.X, pady=(0, 10))

        suggested_questions = [
            "Statut des badges",
            "Liste des badges expirés",
            "Badges expirant ce mois",
            "Nombre de VIP",
            "Statistiques complètes"
        ]

        for question in suggested_questions:
            btn = ttk.Button(suggested_frame, text=question,
                             command=lambda q=question: self.insert_question(q))
            btn.pack(side=tk.TOP, fill=tk.X, pady=2)

        self.chat_history = scrolledtext.ScrolledText(chat_frame, wrap=tk.WORD,
                                                      font=('Helvetica', 10), state='disabled')
        self.chat_history.pack(fill=tk.BOTH, expand=True)

        input_frame = ttk.Frame(chat_frame)
        input_frame.pack(fill=tk.X, pady=5)

        self.user_input = ttk.Entry(input_frame)
        self.user_input.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        send_btn = ttk.Button(input_frame, text="Envoyer", command=self.send_ai_message)
        send_btn.pack(side=tk.RIGHT)

        self.add_to_chat("Assistant: Bonjour! Je suis l'assistant ANP pour la gestion des badges.\n"
                         "Je peux vous fournir:\n"
                         "- Des statistiques sur les badges\n"
                         "- Des listes de badges spécifiques\n"
                         "- Des informations détaillées\n\n"
                         "Posez-moi votre question ou cliquez sur une suggestion ci-dessus.")

    def insert_question(self, question):
        self.user_input.delete(0, tk.END)
        self.user_input.insert(0, question)

    def add_to_chat(self, message):
        self.chat_history.config(state='normal')
        self.chat_history.insert(tk.END, message + "\n\n")
        self.chat_history.config(state='disabled')
        self.chat_history.see(tk.END)

    def send_ai_message(self):
        user_msg = self.user_input.get()
        if not user_msg.strip():
            return

        self.add_to_chat(f"Vous: {user_msg}")
        self.user_input.delete(0, tk.END)

        threading.Thread(target=self.process_ai_response, args=(user_msg,), daemon=True).start()
    
    def compute_status_masks(df):
        import pytz
        import pandas as pd

        df = df.copy()
        today = pd.Timestamp.now(tz=pytz.UTC)

        # Conversion Deactivation_Date
        if 'Deactivation_Date' in df.columns:
            df['Deactivation_Date'] = pd.to_datetime(df['Deactivation_Date'], errors='coerce')
        else:
            df['Deactivation_Date'] = pd.NaT

        # Expirés : date de désactivation passée
        expired_mask = df['Deactivation_Date'].notna() & (df['Deactivation_Date'] < today)

        # Actifs : Token_Status = 1 ou 4 et pas expiré
        active_mask = df['Token_Status'].isin([1,4]) & (~expired_mask)

        # Inactifs : tout le reste
        inactive_mask = ~active_mask

        return active_mask, inactive_mask, expired_mask

# ... (tout le code précédent reste inchangé jusqu'à la méthode generate_ai_response)

    def generate_ai_response(self, user_msg):
        try:
            if not hasattr(self, "df") or self.df is None:
                return "Erreur : les données ne sont pas chargées. Veuillez importer un fichier."

            user_msg = user_msg.lower().strip()
            df = self.df.copy()
            today = datetime.now(pytz.utc)
            response = ""

            if 'Deactivation_Date' in df.columns:
                df['Deactivation_Date'] = pd.to_datetime(df['Deactivation_Date'], errors='coerce')
                days_left = (df['Deactivation_Date'] - today).dt.days
                expired_mask = (days_left < 0)
            else:
                days_left = None
                expired_mask = pd.Series(False, index=df.index)

            # Nouveau calcul des statuts
            active_mask = (df['Token_Status'] == 1) & (~expired_mask)
            inactive_mask = (df['Token_Status'] != 1) | expired_mask
            active_count = active_mask.sum()
            inactive_count = inactive_mask.sum()
            expired_count = expired_mask.sum()

            # Greeting
            if "bonjour" in user_msg or "salut" in user_msg:
                response = (
                    f"Bonjour !\n"
                    f"- Badges actifs : {active_count}\n"
                    f"- Inactifs : {inactive_count}\n"
                    f"- Total : {len(df)}"
                )

            # Full statistics
            elif "statistiques" in user_msg or "complet" in user_msg:
                response = f"Statistiques globales :\n- Total : {len(df)}\n- Actifs : {active_count}\n- Inactifs : {inactive_count}\n"
                if "VIP" in df.columns:
                    response += f"- VIP : {len(df[df['VIP'] == 1])}\n"
                if days_left is not None:
                    response += f"- Expirés : {expired_count}\n"
                    response += f"- Expirant dans 30 jours : {len(df[days_left.between(0, 30)])}"

            # Liste des badges expirés
            elif "liste des badges expirés" in user_msg:
                if days_left is None:
                    return "Les données ne contiennent pas les dates d'expiration."

                expired = df[expired_mask]
                if expired.empty:
                    response = "Aucun badge n'est expiré pour le moment."
                else:
                    response = f"{len(expired)} badge(s) expirés :\n"
                    for _, row in expired.head(20).iterrows():
                        name = f"{row.get('First_Name', '')} {row.get('Last_Name', '')}".strip()
                        date = row['Deactivation_Date'].strftime('%d/%m/%Y')
                        response += f"- {name} (expiré le {date})\n"
                    if len(expired) > 20:
                        response += f"...et {len(expired) - 20} autres."

            # Badges expirant ce mois
            elif "badge expirant ce mois" in user_msg or "badges expirant ce mois" in user_msg:
                if days_left is None:
                    return "Les données ne contiennent pas les dates d'expiration."

                expiring = df[days_left.between(0, 30)]
                if expiring.empty:
                    response = "Aucun badge n'expire ce mois."
                else:
                    response = f"{len(expiring)} badge(s) expirent ce mois :\n"
                    for _, row in expiring.head(20).iterrows():
                        name = f"{row.get('First_Name', '')} {row.get('Last_Name', '')}".strip()
                        date = row['Deactivation_Date'].strftime('%d/%m/%Y')
                        response += f"- {name} (expire le {date})\n"
                    if len(expiring) > 20:
                        response += f"...et {len(expiring) - 20} autres."

            # VIP
            elif "vip" in user_msg:
                if "VIP" in df.columns:
                    vip = df[df['VIP'] == 1]
                    response = f"Nombre total de VIP : {len(vip)}\n"
                    for _, row in vip.head(5).iterrows():
                        name = f"{row.get('First_Name', '')} {row.get('Last_Name', '')}".strip()
                        status = "Actif" if row['Token_Status'] == 1 and (not expired_mask[row.name] if 'Deactivation_Date' in df.columns else True) else "Inactif"
                        response += f"- {name} ({status})\n"
                else:
                    response = "La colonne VIP est absente des données."

            # Statut simple
            elif "statut" in user_msg:
                response = f"- Actifs : {active_count}\n- Inactifs : {inactive_count}"
                if days_left is not None:
                    response += f"\n- Expirés : {expired_count}\n- Expirant dans 30 jours : {len(df[days_left.between(0, 30)])}"

            # Fallback
            else:
                response = (
                    "Je ne comprends pas votre demande. Essayez par exemple :\n"
                    "- 'Liste des badges expirés'\n"
                    "- 'Badges expirant ce mois'\n"
                    "- 'Nombre de VIP'\n"
                    "- 'Statistiques complètes'"
                )

            return response
# ... (le reste du code reste inchangé)
        except Exception as e:
            return f"Désolé, une erreur est survenue : {str(e)}"
if __name__ == "__main__":
    root = tk.Tk()
    app = EnhancedBadgeApp(root)
    root.mainloop()