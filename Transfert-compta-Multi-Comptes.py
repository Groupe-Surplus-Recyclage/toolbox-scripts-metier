import win32com.client
import customtkinter as ctk
from datetime import datetime
import os
import pythoncom
import tempfile
import subprocess
from pypdf import PdfWriter

# --- CONFIGURATION GLOBALE ---
ACCOUNTS = [
    "comptagaillac@groupesurplus.fr", "comptagaillac@gsr-energy.fr",
    "comptagaillac@gsr-logistics.com", "comptagaillac@gsr-repair.com",
    "comptagaillac@surplusindustries.fr", "comptagaillac@surplusmotos.com",
    "comptagaillac@surplusautos.fr"
]
TARGET_FORWARD = "gsr-manager@capture.eu.getyooz.com"
CATEGORY_NAME = "Transfert Yooz"
BASE_DIR = r"C:\_install\transfert_Yooz"
if not os.path.exists(BASE_DIR):
    try: os.makedirs(BASE_DIR)
    except: BASE_DIR = os.getcwd()

EXCLUSIONS_FILE = os.path.join(BASE_DIR, "exclusions.txt")
LOG_FILE = os.path.join(BASE_DIR, "transfert_history.log")

class AppGSR(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Gestion Yooz - Aide à l'exclusion (v3.7)")
        self.geometry("1150x950")
        ctk.set_appearance_mode("light")
        
        try: self.after(200, lambda: self.iconbitmap("yooz-mail.ico"))
        except: pass
        
        self.current_account = ACCOUNTS[0]
        self.email_data = [] 

        # --- UI : TOP ---
        self.frame_top = ctk.CTkFrame(self)
        self.frame_top.pack(pady=(10, 5), padx=20, fill="x")
        ctk.CTkLabel(self.frame_top, text="Compte :", font=("Arial", 12, "bold")).pack(side="left", padx=15)
        self.combo_accounts = ctk.CTkOptionMenu(self.frame_top, values=ACCOUNTS, command=self.change_account, width=300)
        self.combo_accounts.pack(side="left", padx=10, pady=10)
        self.combo_accounts.set(self.current_account)

        # --- UI : EXCLUSIONS ---
        self.frame_excl = ctk.CTkFrame(self)
        self.frame_excl.pack(pady=5, padx=20, fill="x")
        self.header_excl = ctk.CTkFrame(self.frame_excl, fg_color="transparent")
        self.header_excl.pack(fill="x", padx=10, pady=(5, 0))
        ctk.CTkLabel(self.header_excl, text=f"🚫 EXCLUSIONS (Mémorisées dans : {BASE_DIR})", font=("Arial", 10, "bold"), text_color="#C0392B").pack(side="left")
        ctk.CTkButton(self.header_excl, text="📂 Ouvrir dossier", width=120, height=24, fg_color="#D5D8DC", text_color="black", command=self.open_base_dir).pack(side="right")
        
        self.entry_excl = ctk.CTkEntry(self.frame_excl, placeholder_text="Copiez ici l'adresse mail affichée en rouge sous le sujet du mail...", height=30)
        self.entry_excl.pack(pady=(5, 10), padx=10, fill="x")
        self.load_exclusions()
        self.entry_excl.bind("<FocusOut>", lambda e: self.save_exclusions())

        # --- UI : FILTRES ---
        self.frame_filters = ctk.CTkFrame(self); self.frame_filters.pack(pady=5, padx=20, fill="x")
        self.grid_f = ctk.CTkFrame(self.frame_filters, fg_color="transparent"); self.grid_f.pack(pady=5, padx=10)
        self.f_unread = ctk.BooleanVar(value=True); self.f_no_cat = ctk.BooleanVar(value=True); self.f_pdf = ctk.BooleanVar(value=True)
        self.f_subj_f = ctk.BooleanVar(value=True); self.f_body_f = ctk.BooleanVar(value=True)
        self.f_subj_a = ctk.BooleanVar(value=True); self.f_body_a = ctk.BooleanVar(value=True)

        ctk.CTkCheckBox(self.grid_f, text="Non lus", variable=self.f_unread).grid(row=0, column=0, padx=20)
        ctk.CTkCheckBox(self.grid_f, text="Sans catégorie", variable=self.f_no_cat).grid(row=0, column=1, padx=20)
        ctk.CTkCheckBox(self.grid_f, text="PDF requis", variable=self.f_pdf).grid(row=0, column=2, padx=20)
        
        self.r2 = ctk.CTkFrame(self.grid_f, fg_color="transparent"); self.r2.grid(row=1, column=0, columnspan=3, pady=5)
        ctk.CTkCheckBox(self.r2, text="Facture (Objet)", variable=self.f_subj_f).pack(side="left", padx=10)
        ctk.CTkCheckBox(self.r2, text="Facture (Corps)", variable=self.f_body_f).pack(side="left", padx=10)
        ctk.CTkCheckBox(self.r2, text="Avoir (Objet)", variable=self.f_subj_a).pack(side="left", padx=10)
        ctk.CTkCheckBox(self.r2, text="Avoir (Corps)", variable=self.f_body_a).pack(side="left", padx=10)

        # --- UI : CTRL ---
        self.frame_c = ctk.CTkFrame(self, fg_color="transparent"); self.frame_c.pack(pady=5, padx=20, fill="x")
        self.check_auto = ctk.CTkCheckBox(self.frame_c, text="Auto (5 min)"); self.check_auto.pack(side="left", padx=10)
        self.check_read = ctk.CTkCheckBox(self.frame_c, text="Marquer lu"); self.check_read.select(); self.check_read.pack(side="left", padx=10)
        self.check_merge = ctk.CTkCheckBox(self.frame_c, text="Fusionner PDF", text_color="#1F618D"); self.check_merge.pack(side="left", padx=10)
        ctk.CTkButton(self.frame_c, text="🔍 Analyser", fg_color="#5D6D7E", command=self.manual_scan).pack(side="right", padx=15)

        self.scroll_frame = ctk.CTkScrollableFrame(self, width=1100, height=350); self.scroll_frame.pack(pady=10, padx=20, fill="both", expand=True)
        self.frame_buttons = ctk.CTkFrame(self); self.frame_buttons.pack(pady=(0, 10), padx=20, fill="x")
        self.grid_buttons = ctk.CTkFrame(self.frame_buttons, fg_color="transparent"); self.grid_buttons.pack(pady=5)
        
        self.doc_types = [("Facture d'achat", "#3498db"), ("facture sur commande", "#2ecc71"), ("avoir", "#f1c40f"), ("avoir de vente", "#e67e22"), ("facture de vente", "#34495e"), ("credit customer", "#9b59b6"), ("invoice on order", "#16a085"), ("sales invoice", "#2980b9"), ("factura de venta", "#c0392b"), ("autre document", "#95a5a6")]
        self.btns = []
        r, c = 0, 0
        for lbl, clr in self.doc_types:
            b = ctk.CTkButton(self.grid_buttons, text=lbl, fg_color=clr, width=190, height=32, command=lambda l=lbl: self.process_transfer(l), state="disabled")
            b.grid(row=r, column=c, padx=5, pady=3); self.btns.append(b)
            c += 1
            if c > 4: c = 0; r += 1

        self.txt_log = ctk.CTkTextbox(self, height=80, font=("Consolas", 10)); self.txt_log.pack(pady=(0, 10), padx=20, fill="x")
        self.after(1000, self.manual_scan); self.auto_loop()

    def open_base_dir(self):
        if os.path.exists(BASE_DIR): os.startfile(BASE_DIR)

    def load_exclusions(self):
        if os.path.exists(EXCLUSIONS_FILE):
            try:
                with open(EXCLUSIONS_FILE, "r", encoding="utf-8") as f:
                    self.entry_excl.insert(0, f.read().strip())
            except: pass

    def save_exclusions(self):
        try:
            with open(EXCLUSIONS_FILE, "w", encoding="utf-8") as f:
                f.write(self.entry_excl.get())
            self.log("Exclusions sauvegardées.")
        except: pass

    def log(self, m):
        ts = datetime.now().strftime("%H:%M:%S")
        self.txt_log.insert("end", f"[{ts}] {m}\n"); self.txt_log.see("end")

    def change_account(self, n): self.current_account = n; self.manual_scan()

    def get_sender_address(self, msg):
        """Récupère l'adresse mail réelle, même en environnement Exchange/O365."""
        try:
            if msg.SenderEmailType == "EX":
                return msg.Sender.GetExchangeUser().PrimarySmtpAddress.lower()
            return msg.SenderEmailAddress.lower()
        except:
            return str(msg.SenderEmailAddress).lower()

    def scan_folder_recursive(self, folder, results):
        raw_excl = self.entry_excl.get().replace(";", ",").lower()
        excl_list = [e.strip() for e in raw_excl.split(",") if e.strip()]
        
        try:
            items = folder.Items
            query = "[Unread] = true" if self.f_unread.get() else ""
            messages = items.Restrict(query) if query else items
            
            for msg in messages:
                try:
                    sender_addr = self.get_sender_address(msg)
                    sender_name = str(msg.SenderName).lower()
                    
                    # Vérification des exclusions
                    is_excluded = False
                    for term in excl_list:
                        if term in sender_addr or term in sender_name:
                            is_excluded = True
                            break
                    
                    if is_excluded: continue

                    if self.f_no_cat.get() and CATEGORY_NAME in str(msg.Categories): continue
                    pdfs = [a.FileName for a in msg.Attachments if a.FileName.lower().endswith(".pdf")] if msg.Attachments.Count > 0 else []
                    if self.f_pdf.get() and not pdfs: continue
                    
                    subj, body = (str(msg.Subject).lower() if msg.Subject else ""), (str(msg.Body).lower() if msg.Body else "")
                    match = False
                    if self.f_subj_f.get() and "facture" in subj: match = True
                    if self.f_body_f.get() and "facture" in body: match = True
                    if self.f_subj_a.get() and "avoir" in subj: match = True
                    if self.f_body_a.get() and "avoir" in body: match = True
                    if not any([self.f_subj_f.get(), self.f_body_f.get(), self.f_subj_a.get(), self.f_body_a.get()]): match = True
                    
                    if match and pdfs:
                        results.append({'msg': msg, 'pdfs': pdfs, 'folder_name': folder.Name, 'sender': sender_addr})
                    if len(results) > 60: return 
                except: continue

            if folder.Folders.Count > 0:
                for i in range(1, folder.Folders.Count + 1):
                    sub = folder.Folders.Item(i)
                    if sub.Name not in ["Éléments supprimés", "Courrier indésirable", "Deleted Items", "Junk Email", "Sync Issues"]:
                        self.scan_folder_recursive(sub, results)
        except: pass

    def scan_emails(self):
        try:
            pythoncom.CoInitialize()
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            recipient = outlook.CreateRecipient(self.current_account); recipient.Resolve()
            if recipient.Resolved:
                shared_inbox = outlook.GetSharedDefaultFolder(recipient, 6)
                root_folder = shared_inbox.Parent 
                self.log(f"Scan : {self.current_account}...")
                all_results = []
                self.scan_folder_recursive(root_folder, all_results)
                return all_results
            return []
        except Exception as e: self.log(f"Erreur : {e}"); return []

    def manual_scan(self):
        for w in self.scroll_frame.winfo_children(): w.destroy()
        scanned = self.scan_emails(); self.email_data = []
        if scanned:
            for item in scanned:
                m = item['msg']
                f = ctk.CTkFrame(self.scroll_frame, fg_color="transparent"); f.pack(fill="x", pady=5)
                ev = ctk.BooleanVar(value=True); pvs = []
                def toggle(mv=ev, ps=pvs):
                    for _, pv in ps: pv.set(mv.get())
                
                # Checkbox Sujet
                ctk.CTkCheckBox(f, text=f"✉️ ({item['folder_name']}) {m.Subject[:70]}", font=("Arial", 11, "bold"), variable=ev, command=toggle).pack(anchor="w")
                
                # Affichage de l'expéditeur (NOUVEAU)
                ctk.CTkLabel(f, text=f"      Expéditeur : {item['sender']}", font=("Arial", 10, "italic"), text_color="#E74C3C").pack(anchor="w")
                
                for p in item['pdfs']:
                    pv = ctk.BooleanVar(value=True); ctk.CTkCheckBox(f, text=f"      📄 {p}", font=("Arial", 10), variable=pv, text_color="#555").pack(anchor="w", padx=20)
                    pvs.append((p, pv))
                self.email_data.append({'msg': m, 'email_var': ev, 'pdf_vars': pvs})
            for b in self.btns: b.configure(state="normal")
            self.log(f"Trouvé : {len(scanned)} mail(s).")
        else:
            self.log("Aucun mail trouvé."); [b.configure(state="disabled") for b in self.btns]

    def process_transfer(self, prefix):
        merge = self.check_merge.get(); success = 0
        try:
            pythoncom.CoInitialize(); outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            recipient = outlook.CreateRecipient(self.current_account); recipient.Resolve()
            shared_inbox = outlook.GetSharedDefaultFolder(recipient, 6)
            t_root = self.get_or_create_folder(shared_inbox, "FOURNISSEURS")
            t_folder = self.get_or_create_folder(t_root, "AVOIR" if "avoir" in prefix.lower() else "FACTURES")
            for data in self.email_data:
                if data['email_var'].get():
                    msg = data['msg']; sel_pdfs = [pn for pn, pv in data['pdf_vars'] if pv.get()]
                    if not sel_pdfs: continue
                    fwd = msg.Forward(); fwd.To = TARGET_FORWARD; fwd.Subject = f"{prefix} - {msg.Subject}"
                    with tempfile.TemporaryDirectory() as tmp:
                        paths = []
                        for att in msg.Attachments:
                            if att.FileName in sel_pdfs:
                                p = os.path.join(tmp, att.FileName); att.SaveAsFile(p); paths.append(p)
                        if merge and len(paths) > 1:
                            merger = PdfWriter(); mp = os.path.join(tmp, "Fusion_Yooz.pdf")
                            for p in paths: merger.append(p)
                            merger.write(mp); merger.close()
                            while fwd.Attachments.Count > 0: fwd.Attachments.Remove(1)
                            fwd.Attachments.Add(mp)
                        else:
                            while fwd.Attachments.Count > 0: fwd.Attachments.Remove(1)
                            for p in paths: fwd.Attachments.Add(p)
                        fwd.Send(); msg.Categories = CATEGORY_NAME
                        if self.check_read.get(): msg.UnRead = False
                        msg.Save(); msg.Move(t_folder); success += 1
                        ts_h = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                        with open(LOG_FILE, "a", encoding="utf-8") as f_h: f_h.write(f"[{ts_h}] {prefix} | {msg.Subject}\n")
            self.log(f"Succès : {success} transféré(s)."); self.manual_scan()
        except Exception as e: self.log(f"Erreur : {e}")

    def get_or_create_folder(self, p, n):
        try: return p.Folders[n]
        except: return p.Folders.Add(n)

    def auto_loop(self):
        if self.check_auto.get(): self.manual_scan()
        self.after(300000, self.auto_loop)

if __name__ == "__main__":
    app = AppGSR()
    app.mainloop()