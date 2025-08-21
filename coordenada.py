import tkinter as tk
from tkinter import messagebox
import pyautogui
import mouse
import json
import time

class MouseCoordinateTracker:
    def __init__(self, root):
        self.root = root
        self.root.title("Rastreador de Coordenadas (Clique ESQUERDO para Capturar)")
        largura_root = 500
        altura_root = 350  # Aumentado para acomodar os novos botões
        largura_tela = root.winfo_screenwidth()
        root.geometry(f"{largura_root}x{altura_root}+{largura_tela - largura_root}+0")
        
        # Configurações
        self.last_capture_time = 0
        self.capture_delay = 0.5  # 500ms entre capturas
        self.mouse_hook_id = None
        self.app_window_x = root.winfo_x()
        self.app_window_y = root.winfo_y()
        self.app_window_width = 500
        self.app_window_height = 350
        
        self.keys = [
            "pf", "contador_btn", "ok_btn", "cliente", "seguinte", "processamento", "processamento_seguinte", "imprimir", "cancelar", "processamento_sim", "registado", "registado_seguinte",
            "registado_cmp1", "registado_cmp_potencia", "propriedade", "gis_x", "gis_y","estado_instalacao", "num_luz", "quartos", "registro", "registado_sim", "anterior", "fim_rolagem"
        ]
        
        self.current_key_index = 0
        self.coordinates = {}
        self.tracking_active = False

        self.area_click_stage = 0
        self.area_start_point = None

        
        # Frame principal
        self.main_frame = tk.Frame(self.root, bg="#f0f0f0")
        self.main_frame.pack(expand=True, fill=tk.BOTH)
        
        self.setup_ui()
        self.load_existing_coordinates()
        if self.current_key_index < len(self.keys):
            self.start_tracking()
        
        # Atualiza as coordenadas da janela periodicamente
        self.update_window_coords()
    
    def update_window_coords(self):
        self.app_window_x = self.root.winfo_x()
        self.app_window_y = self.root.winfo_y()
        self.app_window_width = self.root.winfo_width()
        self.app_window_height = self.root.winfo_height()
        self.root.after(100, self.update_window_coords)
    
    def setup_ui(self):
        # Limpa o frame principal
        for widget in self.main_frame.winfo_children():
            widget.destroy()
        
        # Configuração do tema
        title_font = ("Arial", 14, "bold")
        label_font = ("Arial", 10)
        
        # Instruções
        self.instructions_label = tk.Label(
            self.main_frame, 
            text=f"▶ Posicione o mouse FORA desta janela e CLIQUE COM BOTÃO ESQUERDO em: {self.keys[self.current_key_index] if self.current_key_index < len(self.keys) else 'COMPLETO'}",
            font=title_font,
            bg="#f0f0f0",
            fg="#0066cc",
            wraplength=450
        )
        self.instructions_label.pack(pady=15)
        
        # Status
        self.status_label = tk.Label(
            self.main_frame,
            text="Aguardando clique ESQUERDO FORA da janela..." if self.current_key_index < len(self.keys) else "Captura concluída!",
            font=label_font,
            bg="#f0f0f0",
            fg="#009933" if self.current_key_index < len(self.keys) else "#006600"
        )
        self.status_label.pack(pady=5)
        
        # Coordenadas
        coord_frame = tk.Frame(self.main_frame, bg="#f0f0f0")
        coord_frame.pack(pady=10)
        
        tk.Label(coord_frame, text="Posição atual:", font=label_font, bg="#f0f0f0").pack(side=tk.LEFT)
        self.current_coord_label = tk.Label(
            coord_frame,
            text="X: 0 | Y: 0",
            font=(label_font[0], 11, "bold"),
            bg="#f0f0f0",
            fg="#333333"
        )
        self.current_coord_label.pack(side=tk.LEFT, padx=10)
        
        # Progresso
        self.progress_label = tk.Label(
            self.main_frame,
            text=f"Progresso: {self.current_key_index + 1}/{len(self.keys)}",
            font=label_font,
            bg="#f0f0f0"
        )
        self.progress_label.pack(pady=10)
        
        # Botões
        btn_frame = tk.Frame(self.main_frame, bg="#f0f0f0")
        btn_frame.pack(pady=15)
        
        self.restart_btn = tk.Button(
            btn_frame,
            text="↻ Reiniciar Tudo",
            command=self.restart_process,
            state=tk.NORMAL if self.current_key_index == len(self.keys) else tk.DISABLED,
            bg="#ff6666",
            fg="white",
            activebackground="#ff3333"
        )
        self.restart_btn.pack(side=tk.LEFT, padx=10)
        
        self.pause_btn = tk.Button(
            btn_frame,
            text="▶ Iniciar" if not self.tracking_active else "⏸ Pausar",
            command=self.toggle_pause,
            bg="#6666ff",
            fg="white",
            activebackground="#3333ff",
            state=tk.NORMAL if self.current_key_index < len(self.keys) else tk.DISABLED
        )
        self.pause_btn.pack(side=tk.LEFT)
    
    def is_click_outside_app(self, x, y):
        """Verifica se o clique foi fora da janela do aplicativo"""
        return not (
            self.app_window_x <= x <= self.app_window_x + self.app_window_width and
            self.app_window_y <= y <= self.app_window_y + self.app_window_height
        )
    
    def start_tracking(self):
        if self.mouse_hook_id is not None:
            mouse.unhook(self.mouse_hook_id)
        
        self.tracking_active = True
        self.status_label.config(text="Aguardando clique ESQUERDO FORA da janela...", fg="#009933")
        self.mouse_hook_id = mouse.hook(self.handle_click)
        self.update_coords()
        self.pause_btn.config(text="⏸ Pausar", state=tk.NORMAL)
    
    def stop_tracking(self):
        self.tracking_active = False
        if self.mouse_hook_id is not None:
            try:
                mouse.unhook(self.mouse_hook_id)
            except ValueError:
                pass
            self.mouse_hook_id = None
        self.status_label.config(text="Captura pausada", fg="#cc0000")
        self.pause_btn.config(text="▶ Continuar", state=tk.NORMAL)
    
    def toggle_pause(self):
        if self.tracking_active:
            self.stop_tracking()
        else:
            self.start_tracking()
    
    def update_coords(self):
        if self.tracking_active:
            x, y = pyautogui.position()
            self.current_coord_label.config(text=f'X: {x} | Y: {y}')
            self.root.after(50, self.update_coords)
    
    def handle_click(self, event):
        if not self.tracking_active:
            return
            
        current_time = time.time()
        
        # Verifica se é um clique ESQUERDO válido e fora da janela
        if (isinstance(event, mouse.ButtonEvent) and 
            event.event_type == "down" and 
            event.button == "left"):
            
            # Obtém a posição atual do mouse
            x, y = pyautogui.position()
            
            if not self.is_click_outside_app(x, y):
                # Clique dentro da janela - não conta como captura
                self.status_label.config(text="✖ Clique FORA da janela para capturar!", fg="#cc0000")
                self.root.after(1000, lambda: self.status_label.config(
                    text="Aguardando clique ESQUERDO FORA da janela...", 
                    fg="#009933"
                ) if self.tracking_active else None)
                return
            
            if (current_time - self.last_capture_time) < self.capture_delay:
                return  # Ignora cliques muito rápidos
            
            self.last_capture_time = current_time
            
            # Feedback visual durante a captura
            self.status_label.config(text="✔ Posição capturada com sucesso!", fg="#009933")
            self.root.update()
            
            current_key = self.keys[self.current_key_index]
            if current_key in ["anterior", "fim_rolagem"]:
            # Se for o primeiro clique
                if "_temp" not in self.coordinates:
                    self.coordinates["_temp"] = (x, y)
                    self.status_label.config(text="✔ Primeiro ponto capturado! Agora clique no canto inferior direito da área.", fg="#0077cc")
                    return
                else:
                    x1, y1 = self.coordinates["_temp"]
                    largura = abs(x - x1)
                    altura = abs(y - y1)
                    x_min = min(x, x1)
                    y_min = min(y, y1)
                    self.coordinates[current_key] = (x_min, y_min, largura, altura
                    )
                    del self.coordinates["_temp"]  # limpa temporário
                    self.status_label.config(text="✔ Área capturada com sucesso!", fg="#009933")
            else:
                self.coordinates[current_key] = (x, y)
                self.status_label.config(text="✔ Posição capturada com sucesso!", fg="#009933")
            
            
            # Atualiza a interface
            self.current_key_index += 1
            self.progress_label.config(text=f"Progresso: {self.current_key_index}/{len(self.keys)}")
            
            if self.current_key_index < len(self.keys):
                self.save_coordinates()
                self.instructions_label.config(
                    text=f"▶ Posicione o mouse FORA desta janela e CLIQUE COM BOTÃO ESQUERDO em: {self.keys[self.current_key_index]}",
                    fg="#0066cc"
                )
                self.status_label.config(text="Aguardando clique ESQUERDO FORA da janela...", fg="#009933")
            else:
                self.finish_capture()
        elif isinstance(event, mouse.ButtonEvent) and event.event_type == "down":
            # Feedback para clique errado
            self.status_label.config(text="✖ Use o BOTÃO ESQUERDO do mouse!", fg="#cc0000")
            self.root.after(1000, lambda: self.status_label.config(
                text="Aguardando clique ESQUERDO FORA da janela...", 
                fg="#009933"
            ) if self.tracking_active else None)
    
    def finish_capture(self):
        self.stop_tracking()
        self.save_coordinates()
        
        self.instructions_label.config(
            text="✅ Todas as coordenadas foram capturadas!",
            fg="#009933"
        )
        self.status_label.config(
            text=f"Arquivo salvo: 'coordenadas.json'",
            fg="#000000"
        )
        self.current_coord_label.config(text="")
        self.restart_btn.config(state=tk.NORMAL)
        self.pause_btn.config(state=tk.DISABLED)
        
        messagebox.showinfo(
            "Concluído", 
            f"Todas as {len(self.keys)} coordenadas foram salvas com sucesso!"
        )
    
    def save_coordinates(self):
        with open("coordenadas.json", "w", encoding='utf-8') as f:
            json.dump(self.coordinates, f, indent=4, ensure_ascii=False)
    
    def load_existing_coordinates(self):
        try:
            with open("coordenadas.json", "r", encoding='utf-8') as f:
                self.coordinates = json.load(f)
                
                # Encontra até onde chegou
                for i, key in enumerate(self.keys):
                    if key not in self.coordinates:
                        self.current_key_index = i
                        break
                else:
                    self.current_key_index = len(self.keys)
                
                if self.current_key_index == len(self.keys):
                    self.finish_capture()
                else:
                    self.instructions_label.config(
                        text=f"▶ Posicione o mouse FORA desta janela e CLIQUE COM BOTÃO ESQUERDO em: {self.keys[self.current_key_index]}",
                        fg="#0066cc"
                    )
                    self.progress_label.config(
                        text=f"Progresso: {self.current_key_index + 1}/{len(self.keys)}"
                    )
        except FileNotFoundError:
            pass
    
    def restart_process(self):
        if self.mouse_hook_id is not None:
            mouse.unhook(self.mouse_hook_id)
            self.mouse_hook_id = None
            
        self.current_key_index = 0
        self.coordinates = {}
        self.tracking_active = False
        self.setup_ui()
        self.restart_btn.config(state=tk.DISABLED)
        self.start_tracking()

if __name__ == "__main__":
    root = tk.Tk()
    app = MouseCoordinateTracker(root)
    root.attributes("-topmost", True)
    root.mainloop()