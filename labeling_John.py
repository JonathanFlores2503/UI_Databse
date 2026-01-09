import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import os
import glob


class VideoAnnotator:
    def __init__(self, root):
        self.root = root
        self.root.title("Anotador PhD - Jonathan Flores - Debug Mode")
        self.root.geometry("1400x900")
        self.root.configure(bg="#f4f4f4")

        self.video_name = ""
        self.frame_paths = []
        self.current_idx = 0
        self.anomalies = [] # [clase, inicio, fin]
        
        self.setup_ui()
        self.bind_keys()

    def setup_ui(self):
        # --- PANEL SUPERIOR ---
        self.top_frame = tk.Frame(self.root, bg="#222", pady=10)
        self.top_frame.pack(fill=tk.X)
        self.sm_canvases = []
        self.sm_labels = []
        for i in range(3):
            f = tk.Frame(self.top_frame, bg="#222")
            f.pack(side=tk.LEFT, expand=True)
            border = 4 if i == 1 else 1
            color = "red" if i == 1 else "white"
            c = tk.Canvas(f, width=200, height=150, bg="black", highlightthickness=border, highlightbackground=color)
            c.pack()
            l = tk.Label(f, text="-", fg="white", bg="#222")
            l.pack()
            self.sm_canvases.append(c)
            self.sm_labels.append(l)

        # --- PANEL CENTRAL ---
        self.mid_panel = tk.Frame(self.root, bg="#f4f4f4")
        self.mid_panel.pack(expand=True, fill=tk.BOTH, padx=20, pady=10)
        self.mid_panel.bind("<Button-1>", lambda e: self.root.focus())

        self.canvas_big = tk.Canvas(self.mid_panel, width=850, height=600, bg="black")
        self.canvas_big.pack(side=tk.LEFT)

        # --- PANEL DE CONTROL ---
        self.ctrl = tk.Frame(self.mid_panel, bg="white", padx=15, pady=15, bd=1, relief=tk.SOLID)
        self.ctrl.pack(side=tk.RIGHT, fill=tk.Y)

        tk.Button(self.ctrl, text="Cargar Video", command=self.load_video, bg="#333", fg="white").pack(fill=tk.X, pady=5)
        
        tk.Label(self.ctrl, text="Clase:", bg="white").pack(anchor=tk.W)
        self.entry_class = tk.Entry(self.ctrl, font=("Arial", 12))
        self.entry_class.insert(0, "Pelea")
        self.entry_class.pack(fill=tk.X, pady=5)

        tk.Label(self.ctrl, text="N Frames Previos:", bg="white").pack(anchor=tk.W, pady=(10,0))
        self.bbox_prev = tk.Entry(self.ctrl, font=("Arial", 12))
        self.bbox_prev.insert(0, "100")
        self.bbox_prev.pack(fill=tk.X)

        tk.Label(self.ctrl, text="N Frames Posteriores:", bg="white").pack(anchor=tk.W, pady=(10,0))
        self.bbox_post = tk.Entry(self.ctrl, font=("Arial", 12))
        self.bbox_post.insert(0, "100")
        self.bbox_post.pack(fill=tk.X, pady=5)

        tk.Label(self.ctrl, text="ACCIONES DE ETIQUETADO", font=("Arial", 9, "bold"), bg="white").pack(pady=(15,0))
        tk.Button(self.ctrl, text="Etiquetar N Previos", command=self.tag_prev, bg="#bbdefb").pack(fill=tk.X, pady=2)
        tk.Button(self.ctrl, text="Etiquetar Solo Actual (S)", command=self.tag_current, bg="#fff9c4").pack(fill=tk.X, pady=2)
        tk.Button(self.ctrl, text="Etiquetar N Posteriores", command=self.tag_post, bg="#c8e6c9").pack(fill=tk.X, pady=2)
        
        tk.Label(self.ctrl, text="SISTEMA", font=("Arial", 9, "bold"), bg="white").pack(pady=(20,0))
        tk.Button(self.ctrl, text="Borrar Última (D)", command=self.delete_tag).pack(fill=tk.X, pady=2)
        tk.Button(self.ctrl, text="Borrar Todas (P)", command=self.clear_all).pack(fill=tk.X, pady=2)
        tk.Button(self.ctrl, text="GUARDAR (G)", command=self.save_txt, bg="#2e7d32", fg="white", font=("Arial", 11, "bold")).pack(fill=tk.X, pady=15)

        self.scroll = tk.Scale(self.root, from_=0, to=100, orient=tk.HORIZONTAL, length=1200, command=self.on_scroll)
        self.scroll.pack(side=tk.BOTTOM, pady=20)

    def bind_keys(self):
        def check_focus(func):
            return lambda e: func() if self.root.focus_get() not in [self.entry_class, self.bbox_prev, self.bbox_post] else None

        self.root.bind("<Right>", check_focus(lambda: self.move(1)))
        self.root.bind("<Left>", check_focus(lambda: self.move(-1)))
        self.root.bind("s", check_focus(self.tag_current))
        self.root.bind("d", check_focus(self.delete_tag))
        self.root.bind("p", check_focus(self.clear_all))
        self.root.bind("g", check_focus(self.save_txt))

    def add_and_merge(self, nueva_clase, inicio, fin):
        print(f"\n--- Intento de etiquetado: {nueva_clase} [{inicio} - {fin}] ---")
        
        mismo_tipo = [a for a in self.anomalies if a[0] == nueva_clase]
        otros_tipos = [a for a in self.anomalies if a[0] != nueva_clase]
        
        mismo_tipo.append([nueva_clase, inicio, fin])
        mismo_tipo.sort(key=lambda x: x[1])
        
        merged = []
        if mismo_tipo:
            curr_clase, curr_start, curr_end = mismo_tipo[0]
            for next_item in mismo_tipo[1:]:
                n_clase, n_start, n_end = next_item
                if n_start <= curr_end + 1: 
                    print(f"DEBUG: Fusionando rango previo [{curr_start}-{curr_end}] con nuevo [{n_start}-{n_end}]")
                    curr_end = max(curr_end, n_end)
                else:
                    merged.append([curr_clase, curr_start, curr_end])
                    curr_start, curr_end = n_start, n_end
            merged.append([curr_clase, curr_start, curr_end])
        
        self.anomalies = otros_tipos + merged
        print(f"RESULTADO ACTUAL: {self.anomalies}")

    def tag_prev(self):
        try:
            n = int(self.bbox_prev.get())
            self.add_and_merge(self.entry_class.get(), max(0, self.current_idx - n), self.current_idx)
        except: print("ERROR: N Previos debe ser un número")

    def tag_current(self):
        self.add_and_merge(self.entry_class.get(), self.current_idx, self.current_idx)

    def tag_post(self):
        try:
            n = int(self.bbox_post.get())
            self.add_and_merge(self.entry_class.get(), self.current_idx, min(len(self.frame_paths)-1, self.current_idx + n))
        except: print("ERROR: N Posteriores debe ser un número")

    def load_video(self):
        folder = filedialog.askdirectory()
        if folder:
            self.video_name = os.path.basename(folder)
            self.frame_paths = sorted(glob.glob(os.path.join(folder, "*.png")) + glob.glob(os.path.join(folder, "*.jpg")))
            if self.frame_paths:
                self.scroll.config(to=len(self.frame_paths)-1)
                self.current_idx = 0
                self.update_view()
                print(f"VIDEO CARGADO: {self.video_name} con {len(self.frame_paths)} frames.")

    def update_view(self):
        if not self.frame_paths: return
        img = Image.open(self.frame_paths[self.current_idx]).resize((850, 600))
        self.photo_big = ImageTk.PhotoImage(img)
        self.canvas_big.create_image(0, 0, anchor=tk.NW, image=self.photo_big)
        for i, offset in enumerate([-1, 0, 1]):
            idx = self.current_idx + offset
            self.sm_canvases[i].delete("all")
            if 0 <= idx < len(self.frame_paths):
                sm = Image.open(self.frame_paths[idx]).resize((200, 150))
                sm_p = ImageTk.PhotoImage(sm)
                setattr(self, f"img_ref_{i}", sm_p)
                self.sm_canvases[i].create_image(0, 0, anchor=tk.NW, image=sm_p)
                self.sm_labels[i].config(text=f"F: {idx}")

    def on_scroll(self, val):
        self.current_idx = int(val)
        self.update_view()

    def move(self, d):
        nv = self.current_idx + d
        if 0 <= nv < len(self.frame_paths): self.scroll.set(nv)

    def delete_tag(self):
        if self.anomalies:
            rem = self.anomalies.pop()
            print(f"ELIMINADO: {rem}")

    def clear_all(self):
        self.anomalies = []
        print("MEMORIA LIMPIA: Todas las etiquetas borradas.")

    def save_txt(self):
        if not self.anomalies: 
            print("AVISO: No hay etiquetas para guardar.")
            return
        line = f"{self.video_name}"
        self.anomalies.sort(key=lambda x: x[1])
        for a in self.anomalies: line += f" {a[0]} {a[1]} {a[2]}"
        with open("Temporal_annotation.txt", "a") as f: 
            f.write(line + "\n")
        print(f"GUARDADO EXITOSO en Temporal_annotation.txt: {line}")
        messagebox.showinfo("Éxito", "Datos guardados y fusionados.")
        self.anomalies = []

if __name__ == "__main__":
    root = tk.Tk()
    app = VideoAnnotator(root)
    root.mainloop()
    print("Fin")