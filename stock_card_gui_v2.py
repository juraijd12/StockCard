from __future__ import annotations

import json
import os
import sqlite3
import subprocess
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from stock_card_backend_v1 import DEFAULT_BATCH, DEFAULT_DB, add_paper, connect


class StockCardApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.project = Path(__file__).resolve().parent
        self.db_path = DEFAULT_DB
        self.config_path = self.project / "settings.json"
        self.settings = self.load_settings()
        self.title("AIYU Warehouse")
        self.geometry("1280x800")
        self.minsize(1000, 650)
        self.configure(bg="#f3f6fb")
        self.style()
        self.build()
        self.refresh_all()

    def load_settings(self):
        data = {"source_folder": "", "new_stockcards_folder": "", "paper_backlog_folder": "", "export_folder": ""}
        if self.config_path.exists():
            try:
                data.update(json.loads(self.config_path.read_text(encoding="utf-8")))
            except (OSError, json.JSONDecodeError):
                pass
        return data

    def save_settings(self):
        self.config_path.write_text(json.dumps(self.settings, indent=2), encoding="utf-8")

    def style(self):
        s = ttk.Style(self)
        try: s.theme_use("clam")
        except tk.TclError: pass
        s.configure("Header.TFrame", background="#172b4d")
        s.configure("Header.TLabel", background="#172b4d", foreground="white", font=("Segoe UI", 19, "bold"))
        s.configure("Sub.TLabel", background="#172b4d", foreground="#b9c9df", font=("Segoe UI", 9))
        s.configure("Panel.TFrame", background="white")
        s.configure("Title.TLabel", background="white", foreground="#1f2937", font=("Segoe UI", 13, "bold"))
        s.configure("Muted.TLabel", background="white", foreground="#64748b", font=("Segoe UI", 9))
        s.configure("Accent.TButton", background="#2563eb", foreground="white", padding=8, font=("Segoe UI", 9, "bold"))
        s.configure("Treeview", rowheight=29, font=("Segoe UI", 9), borderwidth=1, relief="solid")
        s.configure("Treeview.Heading", background="#e8eef7", foreground="#334155", font=("Segoe UI", 9, "bold"), padding=7)

    def build(self):
        header = ttk.Frame(self, style="Header.TFrame", padding=(24, 16)); header.pack(fill="x")
        ttk.Label(header, text="AIYU Warehouse", style="Header.TLabel").pack(anchor="w")
        ttk.Label(header, text="Read, review, and capture stock-card backlog data", style="Sub.TLabel").pack(anchor="w")
        self.tabs = ttk.Notebook(self); self.tabs.pack(fill="both", expand=True, padx=18, pady=18)
        self.material_tab = ttk.Frame(self.tabs, style="Panel.TFrame", padding=18)
        self.review_tab = ttk.Frame(self.tabs, style="Panel.TFrame", padding=18)
        self.paper_tab = ttk.Frame(self.tabs, style="Panel.TFrame", padding=18)
        self.settings_tab = ttk.Frame(self.tabs, style="Panel.TFrame", padding=18)
        for tab, name in [(self.material_tab,"Inventory"),(self.review_tab,"Warnings"),(self.paper_tab,"Paper Backlog"),(self.settings_tab,"Settings")]: self.tabs.add(tab, text=name)
        self.build_materials(); self.build_review(); self.build_paper(); self.build_settings()
        self.status = ttk.Label(self, text="Ready", background="#e8eef7", foreground="#53657d", padding=8, anchor="w"); self.status.pack(fill="x", side="bottom")

    def card(self, parent, label):
        f = ttk.Frame(parent, style="Panel.TFrame", padding=16); f.pack(side="left", fill="x", expand=True, padx=(0, 12))
        value = ttk.Label(f, text="—", background="white", foreground="#2563eb", font=("Segoe UI", 22, "bold")); value.pack(anchor="w")
        ttk.Label(f, text=label, style="Muted.TLabel").pack(anchor="w"); return value

    def build_dashboard(self):
        ttk.Label(self.dashboard_tab, text="Overview", style="Title.TLabel").pack(anchor="w")
        ttk.Label(self.dashboard_tab, text="Current state of imported Excel data and paper backlog", style="Muted.TLabel").pack(anchor="w", pady=(2, 18))
        row = ttk.Frame(self.dashboard_tab, style="Panel.TFrame"); row.pack(fill="x")
        self.d_materials=self.card(row,"MATERIALS"); self.d_tx=self.card(row,"TRANSACTIONS"); self.d_review=self.card(row,"REVIEW QUEUE"); self.d_paper=self.card(row,"PAPER PENDING")
        ttk.Label(self.dashboard_tab, text="Actions", style="Title.TLabel").pack(anchor="w", pady=(28, 10))
        ttk.Button(self.dashboard_tab, text="Scan / Refresh Excel folder", style="Accent.TButton", command=self.scan).pack(anchor="w")
        ttk.Label(self.dashboard_tab, text="Original workbooks are read-only. Scan results are stored in the project database.", style="Muted.TLabel").pack(anchor="w", pady=8)

    def build_materials(self):
        top=ttk.Frame(self.material_tab, style="Panel.TFrame"); top.pack(fill="x")
        ttk.Label(top,text="Materials",style="Title.TLabel").pack(side="left")
        self.material_search=tk.StringVar(); e=ttk.Entry(top,textvariable=self.material_search,width=42); e.pack(side="left",padx=14); e.bind("<KeyRelease>",lambda _e:self.load_materials())
        ttk.Button(top,text="Refresh from Excel",style="Accent.TButton",command=self.scan).pack(side="right")
        self.material_tree=self.table(self.material_tab,("name","location","count","status"),("ITEM","LOCATION","TRANSACTIONS","STATUS"),(330,190,110,150)); self.material_tree.bind("<<TreeviewSelect>>",self.material_selected); self.material_tree.bind("<Double-1>",self.show_material_detail)

    def build_review(self):
        ttk.Label(self.review_tab,text="Review Queue",style="Title.TLabel").pack(anchor="w")
        ttk.Label(self.review_tab,text="Warnings detected by the validator. Double-click a row to inspect its material.",style="Muted.TLabel").pack(anchor="w",pady=(2,12))
        self.review_tree=self.table(self.review_tab,("item","file","block","issues"),("ITEM","SOURCE FILE","CARD","WARNINGS"),(300,260,100,500)); self.review_tree.bind("<Double-1>",self.review_selected)

    def build_paper(self):
        ttk.Label(self.paper_tab,text="Paper Backlog",style="Title.TLabel").pack(anchor="w")
        ttk.Label(self.paper_tab,text="Entries captured from physical records. They remain pending until reviewed.",style="Muted.TLabel").pack(anchor="w",pady=(2,12))
        self.paper_tree=self.table(self.paper_tab,("date","item","type","ref","qty","status"),("DATE","ITEM","TYPE","REFERENCE","QTY","STATUS"),(100,300,100,150,80,150))
        form=ttk.Frame(self.paper_tab,style="Panel.TFrame",padding=(0,16)); form.pack(fill="x")
        self.paper_vars={k:tk.StringVar() for k in ("date","item_id","type","qty","ref","notes")}; self.paper_vars["type"].set("RECEIVING")
        for col,(label,key,width) in enumerate([("Material ID","item_id",10),("Date","date",12),("Type","type",12),("Quantity","qty",10),("Reference","ref",16),("Notes","notes",28)]):
            ttk.Label(form,text=label,style="Muted.TLabel").grid(row=0,column=col*2,sticky="w",padx=(0,4)); v=self.paper_vars[key]
            if key=="type": ttk.Combobox(form,textvariable=v,values=["RECEIVING","ISSUANCE"],state="readonly",width=10).grid(row=1,column=col*2,columnspan=2,padx=(0,10),sticky="ew")
            else: ttk.Entry(form,textvariable=v,width=width).grid(row=1,column=col*2,columnspan=2,padx=(0,10),sticky="ew")
        ttk.Button(form,text="Add pending entry",style="Accent.TButton",command=self.add_paper).grid(row=1,column=12)

    def build_settings(self):
        ttk.Label(self.settings_tab,text="Settings",style="Title.TLabel").pack(anchor="w")
        ttk.Label(self.settings_tab,text="Configure folders for this computer. Source workbooks are never modified by scanning.",style="Muted.TLabel").pack(anchor="w",pady=(2,16))
        self.setting_vars={k:tk.StringVar(value=self.settings.get(k,"")) for k in self.settings}
        labels={"source_folder":"Old/source stock cards","new_stockcards_folder":"New working stock cards","paper_backlog_folder":"Paper backlog","export_folder":"Exports"}
        for key,label in labels.items():
            row=ttk.Frame(self.settings_tab,style="Panel.TFrame"); row.pack(fill="x",pady=5); ttk.Label(row,text=label,width=25,style="Muted.TLabel").pack(side="left"); ttk.Entry(row,textvariable=self.setting_vars[key]).pack(side="left",fill="x",expand=True,padx=10); ttk.Button(row,text="Browse",command=lambda v=self.setting_vars[key]:self.choose(v)).pack(side="right")
        ttk.Button(self.settings_tab,text="Save settings",style="Accent.TButton",command=self.save_settings_ui).pack(anchor="w",pady=16)

    def table(self,parent,columns,headings,widths):
        f=ttk.Frame(parent,style="Panel.TFrame"); f.pack(fill="both",expand=True); t=ttk.Treeview(f,columns=columns,show="headings");
        for c,h,w in zip(columns,headings,widths): t.heading(c,text=h); t.column(c,width=w,anchor="w")
        t.pack(side="left",fill="both",expand=True); s=ttk.Scrollbar(f,orient="vertical",command=t.yview); s.pack(side="right",fill="y"); t.configure(yscrollcommand=s.set); return t

    def refresh_all(self): self.load_materials(); self.load_review(); self.load_paper()
    def load_dashboard(self):
        try:
            c=connect(self.db_path); self.d_materials.config(text=c.execute("select count(*) from materials").fetchone()[0]); self.d_tx.config(text=c.execute("select count(*) from transactions").fetchone()[0]); self.d_review.config(text=c.execute("select count(*) from issues where status='OPEN'").fetchone()[0]); self.d_paper.config(text=c.execute("select count(*) from transactions where origin='PAPER' and status='PENDING_REVIEW'").fetchone()[0]); c.close()
        except sqlite3.Error: pass
    def load_materials(self):
        term=self.material_search.get().upper(); c=connect(self.db_path); rows=c.execute("select id,name,location,transaction_count,last_status from materials where normalized_name like ? or upper(coalesce(location,'')) like ? order by name",(f"%{term}%",f"%{term}%")).fetchall(); c.close(); self.material_tree.delete(*self.material_tree.get_children());
        for r in rows:self.material_tree.insert("","end",iid=str(r["id"]),values=(r["name"],r["location"] or "[blank]",r["transaction_count"],r["last_status"]))
    def material_selected(self,_e=None):
        s=self.material_tree.selection()
        if s:self.tabs.select(self.review_tab) if False else None
    def show_material_detail(self,_e=None):
        s=self.material_tree.selection()
        if not s:return
        material_id=int(s[0]); c=connect(self.db_path); m=c.execute("select * from materials where id=?",(material_id,)).fetchone(); rows=c.execute("select * from transactions where material_id=? order by date,id",(material_id,)).fetchall(); c.close()
        win=tk.Toplevel(self); win.title(f"{m['name']} — transactions"); win.geometry("1150x460")
        ttk.Label(win,text=m['name'],font=("Segoe UI",14,"bold")).pack(anchor="w",padx=16,pady=(14,2)); ttk.Label(win,text=f"Location: {m['location'] or '[blank]'}  •  Status: {m['last_status']}").pack(anchor="w",padx=16,pady=(0,10))
        frame=ttk.Frame(win,padding=10); frame.pack(fill="both",expand=True); cols=("date","receive","from","issue","to","in","out","balance","status"); t=ttk.Treeview(frame,columns=cols,show="headings")
        heads=("DATE","RECEIVING REF.","RECEIVED FROM","ISSUANCE REF.","ISSUED TO","IN","OUT","BALANCE","ORIGIN / STATUS")
        for col,head in zip(cols,heads):t.heading(col,text=head);t.column(col,width=120,anchor="w")
        for r in rows:
            typ=(r['transaction_type'] or '').upper(); t.insert('', 'end', values=(r['date'],r['reference_number'] if typ=='RECEIVING' else '',r['supplier'] or '',r['reference_number'] if typ=='ISSUANCE' else '',r['project'] or '',r['quantity_in'] or '',r['quantity_out'] or '',r['recorded_balance'] or '',f"{r['origin']} / {r['status']}"))
        t.pack(side="left",fill="both",expand=True); sb=ttk.Scrollbar(frame,orient="vertical",command=t.yview);sb.pack(side="right",fill="y");t.configure(yscrollcommand=sb.set)
    def load_review(self):
        c=connect(self.db_path); rows=c.execute("select i.material_id,m.name,s.source_file,s.block_number,s.side_number,i.issue_code from issues i join materials m on m.id=i.material_id left join sources s on s.id=i.source_id where i.status='OPEN' order by m.name").fetchall(); c.close(); self.review_tree.delete(*self.review_tree.get_children());
        for index, r in enumerate(rows):
            self.review_tree.insert("", "end", iid=f"warning_{index}", values=(r["name"], r["source_file"] or "", f"{r['block_number']} / {r['side_number']}", r["issue_code"]))
    def load_paper(self):
        c=connect(self.db_path); rows=c.execute("select t.id,t.date,m.name,t.transaction_type,t.reference_number,t.quantity,t.status from transactions t join materials m on m.id=t.material_id where t.origin='PAPER' order by t.date desc,t.id desc").fetchall(); c.close(); self.paper_tree.delete(*self.paper_tree.get_children());
        for r in rows:self.paper_tree.insert("","end",values=(r["date"],r["name"],r["transaction_type"],r["reference_number"] or "",r["quantity"],r["status"]))
    def review_selected(self,_e=None):
        s=self.review_tree.selection()
        if s:self.material_search.set(self.review_tree.item(s[0])["values"][0]); self.tabs.select(self.material_tab); self.load_materials()
    def add_paper(self):
        try:add_paper(self.db_path,int(self.paper_vars["item_id"].get()),self.paper_vars["date"].get(),self.paper_vars["type"].get(),float(self.paper_vars["qty"].get()),self.paper_vars["ref"].get() or None,self.paper_vars["notes"].get() or None)
        except Exception as exc: messagebox.showerror("Could not save",str(exc)); return
        self.refresh_all(); messagebox.showinfo("Saved","Paper transaction saved as PENDING_REVIEW.")
    def choose(self,var):
        p=filedialog.askdirectory(parent=self)
        if p:var.set(p)
    def save_settings_ui(self):
        self.settings={k:v.get().strip() for k,v in self.setting_vars.items()}; self.save_settings(); self.status.config(text="Settings saved")
    def scan(self):
        source=self.settings.get("source_folder") or str(self.project)
        try:
            subprocess.run([sys.executable,str(self.project/"stock_card_batch_v1.py"),source,"--output",str(DEFAULT_BATCH)],check=True)
            subprocess.run([sys.executable,str(self.project/"stock_card_backend_v1.py"),"--batch",str(DEFAULT_BATCH),"--db",str(self.db_path),"build"],check=True)
            self.refresh_all(); self.status.config(text="Scan complete")
        except (OSError,subprocess.CalledProcessError) as exc: messagebox.showerror("Scan failed",str(exc))


if __name__ == "__main__": StockCardApp().mainloop()
