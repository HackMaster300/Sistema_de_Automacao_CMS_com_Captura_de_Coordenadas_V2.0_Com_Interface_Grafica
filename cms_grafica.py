import json
import os
import pyautogui
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
from datetime import datetime, timedelta
import re
import subprocess
import pyperclip
import keyboard
import time
import pytesseract
from PIL import Image
import re



class Aplicacao:
    def __init__(self):       
        pytesseract.pytesseract.tesseract_cmd = "/usr/bin/tesseract"
         
        self.FILENAME = "coordenadas.json"
        self.LOG_FILENAME = "automacao_log.txt"
        self.coordenadas = {}
        self.cancelar_captura = False
        self.log_messages = []

        # Primeiro criar a janela principal (mas manter escondida)
        self.janela = tk.Tk()
        self.janela.withdraw()
        
        # Mostrar splash screen
        self.splash = self.mostrar_splash_screen()
        self.start_time = time.time()
        
        # Configurar janela principal
        self.configurar_janela_principal()
        self.janela.after(100, self.carregar_recursos)
        
        self.label_arquivo_txt = tk.Label(self.janela, text="Nenhum arquivo carregado", fg="red")
        self.label_arquivo_txt.pack(pady=10)
#função para mostrar tela de carregamento        
    def mostrar_splash_screen(self):
        splash = tk.Toplevel(self.janela)
        splash.title("Carregando...")
        splash.geometry("400x200")
        splash.overrideredirect(True)
        
        # Centralizar
        largura_tela = splash.winfo_screenwidth()
        altura_tela = splash.winfo_screenheight()
        x = (largura_tela - 400) // 2
        y = (altura_tela - 200) // 2
        splash.geometry(f"400x200+{x}+{y}")
        
        # Estilo
        style = ttk.Style()
        style.configure("TFrame", background="#f0f0f0")
        style.configure("TLabel", background="#f0f0f0", font=("Arial", 10))
        style.configure("Title.TLabel", background="#f0f0f0", font=("Arial", 16, "bold"))
        style.configure("Red.TLabel", background="#f0f0f0", foreground="red")
        
        main_frame = ttk.Frame(splash)
        main_frame.pack(expand=True, fill="both", padx=20, pady=20)
        
        ttk.Label(main_frame, text="Automação PyAutoGUI", style="Title.TLabel").pack(pady=10)
        ttk.Label(main_frame, text="Versão 1.0", style="Red.TLabel").pack()
        
        # Usar Progressbar customizado sem thread interna
        canvas = tk.Canvas(main_frame, width=280, height=20, bg='#f0f0f0', highlightthickness=0)
        canvas.pack(pady=20)
        progress_bar = canvas.create_rectangle(0, 0, 0, 20, fill='#4CAF50', outline='')
        
        def update_progress(step=0):
            width = (step % 100) * 2.8
            canvas.coords(progress_bar, 0, 0, width, 20)
            if hasattr(splash, 'running') and splash.running:
                splash.after(50, update_progress, step + 2.5)
        
        ttk.Label(main_frame, text="Inicializando...").pack()
        
        splash.running = True
        splash.after(0, update_progress)
        splash.update()
        return splash
#função para carregar a página principal
    def configurar_janela_principal(self):
        self.janela.title("Automação com PyAutoGUI")
        # Configurações da janela
        largura_janela = 300
        altura_janela = 300
        largura_tela = self.janela.winfo_screenwidth()
        self.janela.geometry(f"{largura_janela}x{altura_janela}+{largura_tela - largura_janela}+0")
        self.janela.attributes("-topmost", True)
        self.janela.overrideredirect(False)
#função para carregar os botões
    def carregar_recursos(self):
        # Carregar dados
        if not self.carregar_coordenadas():
            messagebox.showwarning("Atenção", "Nenhuma coordenada foi encontrada. Vamos capturá-las agora.")
            self.capturar_todas_as_coordenadas()
        
        # Adicionar componentes
        componentes = [
            ("Iniciar Automação", self.executar_automacao),
            # ("Capturar Novas Coordenadas", self.capturar_todas_as_coordenadas),
            # ("Editar Coordenadas", self.editar_coordenadas),
            ("Visualizar Logs", self.show_logs),
            ("Filtrar Logs", self.filter_logs),
            ("Carregar dados", self.selecionar_arquivo_txt)
        ]
        
        for texto, comando in componentes:
            tk.Button(self.janela, text=texto, command=comando).pack(pady=5)
        
        self.add_log("Aplicação iniciada", "Geral")
        
        # Verificar tempo mínimo
        elapsed = time.time() - self.start_time
        if elapsed < 2:
            self.janela.after(int((2 - elapsed) * 1000), self.finalizar_carregamento)
        else:
            self.finalizar_carregamento()
#função para finalizar a inicialização
    def finalizar_carregamento(self):
    # Parar animação de forma segura
        if hasattr(self.splash, 'running'):
            self.splash.running = False
        
        # Fade-out suave
        for alpha in range(10, 0, -1):
            try:
                self.splash.attributes('-alpha', alpha/10)
                self.splash.update()
                time.sleep(0.03)
            except:
                break
        
        try:
            self.splash.destroy()
        except:
            pass
        
        # Mostrar janela principal
        self.janela.deiconify()
#Função para adicionar mensagens ao log com categoria
    def add_log(self, message, category="Geral"):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = f"[{timestamp}] [{category}] {message}"
        self.log_messages.append((timestamp, category, message))
        with open(self.LOG_FILENAME, "a", encoding='utf-8') as log_file:
            log_file.write(log_entry + "\n")
#função para filtrar logs
    def filter_logs(self):
        filter_window = tk.Toplevel(self.janela)
        filter_window.title("Filtrar Logs")
        filter_window.geometry("500x350")
        
        # Variáveis para os filtros
        start_date = tk.StringVar(value=(datetime.now() - timedelta(days=7)).strftime("%Y-%m-%d"))
        end_date = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        operation_types = {
            "Captura": tk.IntVar(value=1),
            "Automação": tk.IntVar(value=1),
            "Edição": tk.IntVar(value=1),
            "Geral": tk.IntVar(value=1)
        }
        
        # Frame para datas
        date_frame = ttk.LabelFrame(filter_window, text="Filtrar por Data")
        date_frame.pack(pady=10, padx=10, fill="x")
        
        ttk.Label(date_frame, text="De:").grid(row=0, column=0, padx=5)
        ttk.Entry(date_frame, textvariable=start_date).grid(row=0, column=1, padx=5)
        ttk.Label(date_frame, text="Até:").grid(row=0, column=2, padx=5)
        ttk.Entry(date_frame, textvariable=end_date).grid(row=0, column=3, padx=5)
        
        # Frame para tipos de operação
        operation_frame = ttk.LabelFrame(filter_window, text="Filtrar por Tipo de Operação")
        operation_frame.pack(pady=10, padx=10, fill="x")
        
        for i, (cat, var) in enumerate(operation_types.items()):
            cb = ttk.Checkbutton(operation_frame, text=cat, variable=var)
            cb.grid(row=0, column=i, padx=5)
        
        # Botão para aplicar filtros
        ttk.Button(filter_window, text="Aplicar Filtros", 
                command=lambda: self.apply_filters(start_date.get(), end_date.get(), operation_types)).pack(pady=10)
    def apply_filters(self, start_date_str, end_date_str, operation_types_vars):
        try:
            start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
            end_date = datetime.strptime(end_date_str, "%Y-%m-%d") + timedelta(days=1)
        except ValueError:
            messagebox.showerror("Erro", "Formato de data inválido. Use YYYY-MM-DD")
            return
        
        selected_categories = [cat for cat, var in operation_types_vars.items() if var.get() == 1]
        
        log_window = tk.Toplevel(self.janela)
        log_window.title("Logs Filtrados")
        log_window.geometry("800x600")
        
        log_text = scrolledtext.ScrolledText(log_window, wrap=tk.WORD)
        log_text.pack(expand=True, fill="both")
        
        try:
            with open(self.LOG_FILENAME, "r", encoding='utf-8') as log_file:
                for line in log_file:
                    try:
                        match = re.match(r'\[(.*?)\] \[(.*?)\] (.*)', line)
                        if match:
                            log_date_str, category, message = match.groups()
                            log_date = datetime.strptime(log_date_str, "%Y-%m-%d %H:%M:%S")
                            
                            if (start_date <= log_date <= end_date and 
                                category in selected_categories):
                                log_text.insert(tk.END, line)
                    except:
                        continue
        except FileNotFoundError:
            log_text.insert(tk.END, "Nenhum log disponível ainda.")
        
        log_text.config(state=tk.DISABLED)
#Função para mostrar todos os logs
    def show_logs(self):
        log_window = tk.Toplevel(self.janela)
        log_window.title("Logs de Execução")
        log_window.geometry("800x600")
        
        log_text = scrolledtext.ScrolledText(log_window, wrap=tk.WORD)
        log_text.pack(expand=True, fill="both")
        
        try:
            with open(self.LOG_FILENAME, "r", encoding='utf-8') as log_file:
                logs = log_file.read()
            log_text.insert(tk.END, logs)
        except FileNotFoundError:
            log_text.insert(tk.END, "Nenhum log disponível ainda.")
        
        log_text.config(state=tk.DISABLED)
#Função para salvar coordenadas no JSON
    def salvar_coordenadas(self):
        with open(self.FILENAME, "w") as f:
            json.dump(coordenadas, f, indent=4)
        messagebox.showinfo("Salvo", "Coordenadas salvas com sucesso!")
        self.add_log("Coordenadas salvas no arquivo JSON", "Geral")
#Função para capturar coordenadas ao clicar na tela
    def capturar_coordenadas(self, chave):
        global cancelar_captura
        if cancelar_captura:
            return
        
        resposta = messagebox.askyesno("Captura", f"Capturar posição para: {chave}?")
        if not resposta:
            cancelar_captura = True
            self.add_log(f"Captura de {chave} cancelada pelo usuário", "Captura")
            return
        
        messagebox.showinfo("Captura", f"Por favor, clique na posição desejada para: {chave}.")
        self.add_log(f"Iniciando captura para {chave}", "Captura")
        x, y = pyautogui.position()
        coordenadas[chave] = (x, y)
        messagebox.showinfo("Capturado", f"{chave} capturado em ({x}, {y}).")
        self.add_log(f"{chave} capturado nas coordenadas ({x}, {y})", "Captura")
#Função para capturar todas as coordenadas
    def capturar_todas_as_coordenadas(self):
        try:
            self.add_log("Iniciando captura de coordenadas", "Captura") 

            # Caminho para o arquivo de captura (ajuste conforme sua necessidade)
            caminho_captura = "/home/zebito/Downloads/drive/cms/linux/aprendizado.py"  # ← MODIFIQUE AQUI
            
            # Verifica se o arquivo existe
            if not os.path.exists(caminho_captura):
                messagebox.showerror("Erro", f"Arquivo de captura não encontrado em:\n{caminho_captura}")
                self.add_log(f"Arquivo de captura não encontrado: {caminho_captura}", "Captura")
                return
            
            # Executa o arquivo
            resultado = subprocess.run(['python', caminho_captura], 
                                    capture_output=True, text=True)
            
            if resultado.returncode == 0:
                messagebox.showinfo("Sucesso", "Captura de coordenadas concluída!")
                self.add_log("Captura de coordenadas concluída com sucesso", "Captura")
                self.carregar_coordenadas()
            else:
                messagebox.showerror("Erro", f"Falha na captura:\n{resultado.stderr}")
                self.add_log(f"Falha na captura: {resultado.stderr}", "Captura")
                
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível executar a captura:\n{str(e)}")
            self.add_log(f"Erro ao executar captura: {str(e)}", "Captura")            
#Função para carregar coordenadas do JSON
    def carregar_coordenadas(self):
        global coordenadas
        if os.path.exists(self.FILENAME):
            with open(self.FILENAME, "r") as f:
                coordenadas = json.load(f)
            self.add_log("Coordenadas carregadas do arquivo JSON", "Geral")
            return True
        self.add_log("Arquivo de coordenadas não encontrado", "Geral")
        return False
#Função para executar a automação
    def executar_automacao(self):
        try:
            self.add_log("Iniciando execução da automação", "Automação")
            erro_img_path = "/home/zebito/Downloads/drive/cms/linux/erro.png" 
            erro_img_path2 = "/home/zebito/Downloads/drive/cms/linux/erro2.png"
            erro_img_path1 = "/home/zebito/Downloads/drive/cms/linux/erro1.png"
            # Ajuste conforme necessário

            pf = (coordenadas["pf"])
            x_contador_btn, y_contador_btn = 366, 251
            x_ok_btn, y_ok_btn = 846, 475  # Coordenadas do botão "OK" no erro
            x_cliente, y_clique_esquero = 490, 404
            x_seguinte, y_seguinte = 512, 735
            x_processamento,y_processamento = 351, 356
            x_processamento_seguinte, y_processamento_seguinte = 507, 667
            x_imprimir, y_imprimir = 563, 191
            x_processamento_sim, y_processamento_sim = 507, 667
            x_registado, y_registado = 542, 360
            x_registado_seguinte, y_registado_seguinte = 507, 667
            x_registado_cmp1, y_registado_cmp1 = 471, 180                       
            x_registado_cmp_potencia, y_registado_cmp1_potencia = 807, 283
            x_propriedade, y_propriedade = 791, 308
            x_gis_x, y_gis_x = 777, 332
            x_gis_y, y_gis_y = 777, 355
            x_estado_instalacao, y_estado_instalacao = 479, 393
            x_num_luz, y_num_luz = 775, 393
            x_quartos, y_quartos = 775, 412
            x_registro, y_registro = 507, 667
            x_registado_sim, y_registado_sim = 507, 667
            x_anterior, y_anterior = 385, 734
            area2 = (442, 248, 87, 16)
            click_position = (935, 660) 
            x_cancelar, y_cancelar = (616, 483)


            # Definir a área da tela para capturar (x, y, largura, altura)
            total_processados = 0
            contador_em_processamento = 0
            contador_registados = 0
            contador_introduzido = 0
            numeros_errados = 0  
            numero = ""
            numeros = []


            try:
                time.sleep(2)
                for item in dados:
                    pyautogui.click(pf, button='left', clicks=2, interval=0.25)
                    time.sleep(0.9)

                    pyautogui.click(x_contador_btn, y_contador_btn, button='left', clicks=1, interval=0.25)
                    time.sleep(0.1)

                    pyperclip.copy(item)

                    pyautogui.hotkey("ctrl", "v")
                    print(item)
                    time.sleep(1)

                    pyautogui.press("enter")
                    time.sleep(0.8)
                                
                    total_processados += 1

                    try:
                        erro_na_tela = pyautogui.locateOnScreen(erro_img_path, confidence=0.5)
                        erro_na_tela1 = pyautogui.locateOnScreen(erro_img_path1, confidence=0.5)  # Aumentar a confiança

                        if erro_na_tela:
                            numeros_errados += 1
                            numero += item
                            numeros = re.findall(r'\d{11}', numero)
                            print("Erro detectado, clicando em OK...")
                            # Clica no botão "OK"
                            pyautogui.click(x_ok_btn, y_ok_btn, button='left', clicks=1, interval=0.25)   

                            continue        
                
                    except pyautogui.ImageNotFoundException:
                        # print("Erro não encontrado.")
                        # Clica com o botão direito na coordenada especificada
                        pyautogui.click(x_clique_esquerdo, y_clique_esquero, button='left', clicks=2, interval=0.25)
                        time.sleep(2)
                        keyboard.press_and_release("space")
                        time.sleep(0.1)  

                        # Verifica novamente se o erro apareceu após o clique direito
                        try:
                            erro_na_tela2 = pyautogui.locateOnScreen(erro_img_path2, confidence=0.7)
                            if erro_na_tela2:
                                # print("Erro detectado após clique direito, clicando em OK...")
                                # Clica no botão "OK"
                                pyautogui.click(x_ok_btn, y_ok_btn, button='left', clicks=1, interval=0.25)
                                time.sleep(2)

                                # Pressiona Enter novamente
                                keyboard.press_and_release("enter")
                                time.sleep(2)
                                break
                            else:
                                print("Erro não detectado após clique direito.")

                        except pyautogui.ImageNotFoundException:
                            print("Erro não encontrado após clique direito.")
                    
                    keyboard.press_and_release("space")
                    # Segunda parte do código: captura de tela e análise de texto
                    screenshot = pyautogui.screenshot(region=area)

                    # Usar o pytesseract para extrair texto da captura de tela
                    texto_extraido = pytesseract.image_to_string(screenshot)

                    # Mostrar o texto extraído (para depuração)
                    print("Texto extraído:\n", texto_extraido)

                    # Dividir o texto extraído em linhas, removendo espaços em branco extras
                    linhas = [linha.strip() for linha in texto_extraido.split('\n') if linha.strip()]
                    
                    total_grupos = len(linhas) // 3
                    tipos = linhas[:total_grupos]
                    estados = linhas[total_grupos:2 * total_grupos]
                    centros = linhas[2 * total_grupos:]

                    # Combinar os dados correspondentes
                    linhas_combinadas = []
                    for i in range(total_grupos):
                        linha_completa = f"{tipos[i]} {estados[i]} {centros[i]}"
                        linhas_combinadas.append(linha_completa)

                    # Mostrar as linhas combinadas
                    print("\nLinhas Combinadas:")
                    for linha in linhas_combinadas:
                        print(linha)

                    encontrou_linha = False
                    
                    
                    # Garantir que temos um número de linhas múltiplo de 3 (para combinar corretamente)
                    if len(linhas) % 3 != 0:
                        print("Erro: O número de linhas não é múltiplo de 3. Verifique a captura.")
                        
                        # Definir os comportamentos baseados em palavras-chave
                        encontrou_linha = False
                        for i, linha in enumerate(linhas_combinadas):
                            if 'suspeita de fraude' in linha.lower(): 
                                if 'registad' in linha.lower():
                                    print(f"Linha encontrada (Registado): {linha}")
                                    encontrou_linha = True

                                    contador_registados +=1
                                    # Captura os primeiros dígitos (OT) no início da linha
                                    match = re.match(r'^\d+', linha)
                                    if match:
                                        ot = match.group()
                                        print(f"OT encontrado: {ot}")

                                        altura_linha = 20  # Ajuste conforme necessário para a altura da linha real
                                        y_pos = area[1] + (i * altura_linha) + 5
                                        x_pos = area[0] + 30  # Ajuste conforme necessário para a primeira coluna

                                        # Mover o mouse para a posição da linha específica e clicar
                                        pyautogui.moveTo(x_pos, y_pos)
                                        pyautogui.click()
                                        
                                        pyautogui.click(x_seguinte, y_seguinte, button='left', clicks=2, interval=0.25)
                                        time.sleep(1)

                                        pyautogui.click(x_processamento,y_processamento, button='left', clicks=2, interval=0.25)
                                        time.sleep(1)

                                        pyautogui.click(x_processamento_seguinte, y_processamento_seguinte, button='left', clicks=2, interval=0.25)
                                        time.sleep(4)

                                        pyautogui.click(x_imprimir, y_imprimir, button='left', clicks=1, interval=0.25)
                                        time.sleep(2)


                                        pyautogui.click( x_processamento_sim, y_processamento_sim, button='left', clicks=1, interval=0.25)
                                        time.sleep(4)

                                        pyautogui.click(x_cancelar, y_cancelar, button='left', clicks=1, interval=0.25)
                                        time.sleep(2)

                                        pyautogui.click(x_registado, y_registado, button='left', clicks=2, interval=0.25)
                                        time.sleep(2)

                                        pyautogui.click(x_registado_seguinte, y_registado_seguinte, button='left', clicks=2, interval=0.25)
                                        time.sleep(3)

                                        keyboard.press_and_release('space')

                                        pyautogui.click(x_registado_cmp1, y_registado_cmp1 , button='left', clicks=1, interval=0.25)
                                        time.sleep(5)

                                        keyboard.press_and_release('r')
                                        time.sleep(0.5)

                                        keyboard.press_and_release('tab')
                                        time.sleep(0.5)

                                        keyboard.press_and_release('v')
                                        time.sleep(0.5)

                                        keyboard.press_and_release('tab')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('2')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('tab')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('2')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('shift+tab')
                                        time.sleep(2)

                                        screenshot = pyautogui.screenshot(region=area2)
                                        texto_extraido = pytesseract.image_to_string(screenshot)
                                        
                                        texto_extraido = texto_extraido.strip()  # Remove espaços no início e no fim
                                        texto_extraido = texto_extraido.replace("\n", "")  # Remove quebras de linha

                                        # Mostrar o texto extraído para depuração
                                        print("Texto extraído:\n", texto_extraido)

                                        # Comparar o texto
                                        if texto_extraido == "22000 V":
                                            keyboard.press_and_release('tab')
                                            time.sleep(0.3)
                                            keyboard.press_and_release('2')
                                            time.sleep(0.3)
                                            
                                        else:
                                            keyboard.press_and_release('tab')
                                            time.sleep(0.3)
                                            
                                        keyboard.press_and_release('tab')
                                        time.sleep(0.3)
                                        keyboard.press_and_release('s')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('tab')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('s')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('tab')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('r')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('tab')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('2')
                                        keyboard.press_and_release('0')
                                        keyboard.press_and_release('0')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('tab')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('4')
                                        keyboard.press_and_release('3')
                                        time.sleep(0.3)
                    
                                        keyboard.press_and_release('tab')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('4')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('tab')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('2')
                                        keyboard.press_and_release('0')
                                        keyboard.press_and_release('0')
                                        time.sleep(1)
                                        
                                        pyautogui.click(x_registado_cmp_potencia, y_registado_cmp1_potencia, button='left', clicks=1, interval=0.25)
                                        time.sleep(0.3)

                                        keyboard.press_and_release('2')
                                        keyboard.press_and_release('.')
                                        keyboard.press_and_release('2')
                                        time.sleep(0.3)
                                        
                                        pyautogui.click(x_propriedade, y_propriedade, button='left', clicks=2, interval=0.25)
                                        time.sleep(0.3)
                                        
                                        keyboard.press_and_release('backspace')
                                        time.sleep(0.3) 

                                        keyboard.press_and_release('1')
                                        keyboard.press_and_release('7')
                                        keyboard.press_and_release('2')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('tab')
                                        time.sleep(0.3)

                                        pyautogui.click(x_gis_x, y_gis_x, button='left', clicks=2, interval=0.25)
                                        time.sleep(0.3)
                                        
                                        keyboard.press_and_release('backspace')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('2')
                                        time.sleep(0.3)

                                        pyautogui.click(x_gis_y, y_gis_y, button='left', clicks=2, interval=0.25)
                                        time.sleep(0.3)
                                        
                                        keyboard.press_and_release('backspace')
                                        time.sleep(0.3)
                                        
                                        keyboard.press_and_release('1')
                                        time.sleep(0.3)

                                        
                                        pyautogui.click(x_estado_instalacao, y_estado_instalacao, button='left', clicks=1, interval=0.25)
                                        time.sleep(0.3) 

                                        keyboard.press_and_release('s')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('tab')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('s')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('tab')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('s')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('tab')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('s')
                                        time.sleep(0.3)
                                    
                                        pyautogui.click(x_num_luz, y_num_luz , button='left', clicks=1, interval=0.25)
                                        time.sleep(0.3) 

                                        keyboard.press_and_release('8')
                                        time.sleep(1)

                                        pyautogui.click(x_quartos, y_quartos, button='left', clicks=1, interval=0.25)
                                        time.sleep(0.3) 

                                        keyboard.press_and_release('4')
                                        time.sleep(0.3)
                                        

                                        pyautogui.click(x_registro, y_registro, button='left', clicks=1, interval=0.25)
                                        time.sleep(0.3) 

                                        keyboard.press_and_release('space')
                                        keyboard.press_and_release('space')
                                        keyboard.press_and_release('space')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('tab')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('n')
                                        time.sleep(0.3)
                                        

                                        pyautogui.click( x_registado_sim, y_registado_sim, button='left', clicks=1, interval=0.25)
                                        time.sleep(0.3) 

                                        keyboard.press_and_release('space')
                                        keyboard.press_and_release('space')
                                        keyboard.press_and_release('space')
                                        time.sleep(0.3)
                                        
                                        pyautogui.click(x_anterior, y_anterior, button='left', clicks=3, interval=0.25)
                                        time.sleep(0.1)
                                    

                                elif 'em processamento' in linha.lower():
                                    print(f"Linha encontrada (Em processamento): {linha}")
                                    # print(f"Linha encontrada (Em Processamento): {linha}")
                                    encontrou_linha = True
                                    contador_em_processamento += 1

                                    match = re.match(r'^\d+', linha)
                                    if match:
                                        ot = match.group()
                                        print(f"OT encontrado: {ot}")

                                        altura_linha = 20  # Ajuste conforme necessário para a altura da linha real
                                        y_pos = area[1] + (i * altura_linha) + 5
                                        x_pos = area[0] + 30

                                        pyautogui.moveTo(x_pos, y_pos)
                                        pyautogui.click()

                                        pyautogui.click(x_seguinte, y_seguinte, button='left', clicks=2, interval=0.25)
                                        time.sleep(1)

                                        pyautogui.click(x_registado, y_registado, button='left', clicks=2, interval=0.25)
                                        time.sleep(1)

                                        pyautogui.click(x_processamento_seguinte, y_processamento_seguinte, button='left', clicks=1, interval=0.25)
                                        time.sleep(5)

                                        pyautogui.click(x_registado_cmp1, y_registado_cmp1, button='left', clicks=1, interval=0.25)
                                        time.sleep(4)

                                        keyboard.press_and_release('r')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('tab')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('v')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('tab')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('2')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('tab')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('2')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('shift+tab')
                                        time.sleep(2)

                                        screenshot = pyautogui.screenshot(region=area2)
                                        texto_extraido = pytesseract.image_to_string(screenshot)
                                        
                                        texto_extraido = texto_extraido.strip()  # Remove espaços no início e no fim
                                        texto_extraido = texto_extraido.replace("\n", "")  # Remove quebras de linha

                                        # Mostrar o texto extraído para depuração
                                        print("Texto extraído:\n", texto_extraido)

                                        # Comparar o texto
                                        if texto_extraido == "22000 V":
                                            keyboard.press_and_release('tab')
                                            time.sleep(0.3)
                                            keyboard.press_and_release('2')
                                            time.sleep(0.3)
                                            
                                        else:
                                            keyboard.press_and_release('tab')
                                            time.sleep(0.3)
                                            
                                        keyboard.press_and_release('tab')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('s')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('tab')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('s')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('tab')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('r')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('tab')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('2')
                                        keyboard.press_and_release('0')
                                        keyboard.press_and_release('0')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('tab')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('4')
                                        keyboard.press_and_release('3')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('tab')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('4')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('tab')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('2')
                                        keyboard.press_and_release('0')
                                        keyboard.press_and_release('0')
                                        time.sleep(1)
                                        
                                        pyautogui.click(x_registado_cmp_potencia, y_registado_cmp1_potencia, button='left', clicks=1, interval=0.25)
                                        time.sleep(0.3)

                                        keyboard.press_and_release('2')
                                        keyboard.press_and_release('.')
                                        keyboard.press_and_release('2')
                                        time.sleep(0.3)
                                        
                                        pyautogui.click(x_propriedade, y_propriedade, button='left', clicks=2, interval=0.25)
                                        time.sleep(0.3)
                                        
                                        keyboard.press_and_release('backspace')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('1')
                                        keyboard.press_and_release('7')
                                        keyboard.press_and_release('2')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('tab')
                                        time.sleep(0.3)

                                        pyautogui.click(x_gis_x, y_gis_x, button='left', clicks=2, interval=0.25)
                                        time.sleep(0.3)
                                        
                                        keyboard.press_and_release('backspace')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('2')
                                        time.sleep(0.3)

                                        pyautogui.click(x_gis_y, y_gis_y, button='left', clicks=2, interval=0.25)
                                        time.sleep(0.3)
                                        
                                        keyboard.press_and_release('backspace')
                                        time.sleep(0.3)
                                        
                                        keyboard.press_and_release('1')
                                        time.sleep(0.3)

                                        pyautogui.click(x_estado_instalacao, y_estado_instalacao, button='left', clicks=1, interval=0.25)
                                        time.sleep(0.3) 

                                        keyboard.press_and_release('s')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('tab')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('s')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('tab')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('s')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('tab')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('s')
                                        time.sleep(0.3)

                                        pyautogui.click(x_num_luz, y_num_luz, button='left', clicks=1, interval=0.25)
                                        time.sleep(0.3) 

                                        keyboard.press_and_release('8')
                                        time.sleep(1)

                                        pyautogui.click(x_quartos, y_quartos, button='left', clicks=1, interval=0.25)
                                        time.sleep(0.3) 

                                        keyboard.press_and_release('4')
                                        time.sleep(0.3)

                                        pyautogui.click(x_registro, y_registro, button='left', clicks=1, interval=0.25)
                                        time.sleep(0.3) 

                                        keyboard.press_and_release('space')
                                        keyboard.press_and_release('space')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('tab')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('n')
                                        time.sleep(0.3)

                                        pyautogui.click(x_registado_sim, y_registado_sim, button='left', clicks=1, interval=0.25)
                                        time.sleep(0.3) 

                                        keyboard.press_and_release('space')
                                        keyboard.press_and_release('space')
                                        time.sleep(0.3)

                                        pyautogui.click(x_anterior, y_anterior, button='left', clicks=3, interval=0.25)
                                        time.sleep(0.1)
                                        
                                        
                            
                        if not encontrou_linha:
                            pyautogui.click(click_position)
                            time.sleep(0.1)

                            keyboard.press_and_release('page down')
                            time.sleep(0.5)
                            # print("Nenhuma linha correspondente encontrada. Pressionando Ctrl + I.")
                            screenshot = pyautogui.screenshot(region=area)

                # Usar o pytesseract para extrair texto da captura de tela
                            texto_extraido = pytesseract.image_to_string(screenshot)

                            # Mostrar o texto extraído (para depuração)
                            print("Texto extraído:\n", texto_extraido)

                            # Dividir o texto extraído em linhas, removendo espaços em branco extras
                            linhas = [linha.strip() for linha in texto_extraido.split('\n') if linha.strip()]
                        

                            total_grupos = len(linhas) // 3
                            tipos = linhas[:total_grupos]
                            estados = linhas[total_grupos:2 * total_grupos]
                            centros = linhas[2 * total_grupos:]

                            # Combinar os dados correspondentes
                            linhas_combinadas = []
                            for i in range(total_grupos):
                                linha_completa = f"{tipos[i]} {estados[i]} {centros[i]}"
                                linhas_combinadas.append(linha_completa)

                            # Mostrar as linhas combinadas
                            print("\nLinhas Combinadas:")
                            for linha in linhas_combinadas:
                                print(linha)

                            # Definir os comportamentos baseados em palavras-chave
                            encontrou_linha = False

                            for i, linha in enumerate(linhas_combinadas):
                                if 'suspeita de fraude' in linha.lower():
                                    if 'registad' in linha.lower():
                                        print(f"Linha encontrada (Registado): {linha}")
                                        encontrou_linha = True
                                        contador_registados +=1

                                        # Extrair o número da O.T. usando regex (números no início da linha)
                                        match = re.match(r'^\d+', linha)
                                        if match:
                                            ot = match.group()
                                            print(f"Número da O.T.: {ot}")

                                            # Calcular a posição Y para o clique
                                            altura_linha = 20  # Ajuste conforme necessário para a altura da linha real
                                            y_pos = area[1] + (i * altura_linha) + 5
                                            x_pos = area[0] + 30  # Ajuste conforme necessário para a primeira coluna

                                            # Mover o mouse para a posição da linha específica e clicar
                                            pyautogui.moveTo(x_pos, y_pos)
                                            pyautogui.click()
                                            print(f"Mouse movido e clique realizado na linha com O.T.: {ot}")

                                            pyautogui.click(x_seguinte, y_seguinte, button='left', clicks=2, interval=0.25)
                                            time.sleep(1)

                                            pyautogui.click(x_processamento,y_processamento, button='left', clicks=2, interval=0.25)
                                            time.sleep(1)

                                            pyautogui.click(x_processamento_seguinte, y_processamento_seguinte, button='left', clicks=2, interval=0.25)
                                            time.sleep(4)

                                            pyautogui.click(x_imprimir, y_imprimir, button='left', clicks=1, interval=0.25)
                                            time.sleep(2)


                                            pyautogui.click( x_processamento_sim, y_processamento_sim, button='left', clicks=1, interval=0.25)
                                            time.sleep(4)

                                            pyautogui.click(x_cancelar, y_cancelar, button='left', clicks=1, interval=0.25)
                                            time.sleep(2)

                                            pyautogui.click(x_registado, y_registado, button='left', clicks=2, interval=0.25)
                                            time.sleep(1)

                                            pyautogui.click(x_registado_seguinte, y_registado_seguinte, button='left', clicks=2, interval=0.25)
                                            time.sleep(3)
                                            keyboard.press_and_release('space')

                                            pyautogui.click(x_registado_cmp1, y_registado_cmp1 , button='left', clicks=1, interval=0.25)
                                            time.sleep(5)

                                            keyboard.press_and_release('r')
                                            time.sleep(0.5)

                                            keyboard.press_and_release('tab')
                                            time.sleep(0.5)

                                            keyboard.press_and_release('v')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('tab')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('2')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('tab')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('2')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('shift+tab')
                                            time.sleep(2)

                                            screenshot = pyautogui.screenshot(region=area2)
                                            texto_extraido = pytesseract.image_to_string(screenshot)
                                            
                                            texto_extraido = texto_extraido.strip()  # Remove espaços no início e no fim
                                            texto_extraido = texto_extraido.replace("\n", "")  # Remove quebras de linha

                                            # Mostrar o texto extraído para depuração
                                            print("Texto extraído:\n", texto_extraido)

                                            # Comparar o texto
                                            if texto_extraido == "22000 V":
                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)
                                                keyboard.press_and_release('2')
                                                time.sleep(0.3)
                                                
                                            else:
                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)
                                                
                                            keyboard.press_and_release('tab')
                                            time.sleep(0.3)
                                            keyboard.press_and_release('s')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('tab')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('s')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('tab')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('r')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('tab')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('2')
                                            keyboard.press_and_release('0')
                                            keyboard.press_and_release('0')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('tab')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('4')
                                            keyboard.press_and_release('3')
                                            time.sleep(0.3)
                        
                                            keyboard.press_and_release('tab')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('4')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('tab')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('2')
                                            keyboard.press_and_release('0')
                                            keyboard.press_and_release('0')
                                            time.sleep(1)
                                            
                                            pyautogui.click(x_registado_cmp_potencia, y_registado_cmp1_potencia, button='left', clicks=1, interval=0.25)
                                            time.sleep(0.3)

                                            keyboard.press_and_release('2')
                                            keyboard.press_and_release('.')
                                            keyboard.press_and_release('2')
                                            time.sleep(0.3)
                                            
                                            pyautogui.click(x_propriedade, y_propriedade, button='left', clicks=2, interval=0.25)
                                            time.sleep(0.3)
                                            
                                            keyboard.press_and_release('backspace')
                                            time.sleep(0.3) 

                                            keyboard.press_and_release('1')
                                            keyboard.press_and_release('7')
                                            keyboard.press_and_release('2')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('tab')
                                            time.sleep(0.3)

                                            pyautogui.click(x_gis_x, y_gis_x, button='left', clicks=2, interval=0.25)
                                            time.sleep(0.3)
                                            
                                            keyboard.press_and_release('backspace')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('2')
                                            time.sleep(0.3)

                                            pyautogui.click(x_gis_y, y_gis_y, button='left', clicks=2, interval=0.25)
                                            time.sleep(0.3)
                                            
                                            keyboard.press_and_release('backspace')
                                            time.sleep(0.3)
                                            
                                            keyboard.press_and_release('1')
                                            time.sleep(0.3)

                                            
                                            pyautogui.click(x_estado_instalacao, y_estado_instalacao, button='left', clicks=1, interval=0.25)
                                            time.sleep(0.3) 

                                            keyboard.press_and_release('s')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('tab')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('s')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('tab')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('s')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('tab')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('s')
                                            time.sleep(0.3)
                                        
                                            pyautogui.click(x_num_luz, y_num_luz , button='left', clicks=1, interval=0.25)
                                            time.sleep(0.3) 

                                            keyboard.press_and_release('8')
                                            time.sleep(1)

                                            pyautogui.click(x_quartos, y_quartos, button='left', clicks=1, interval=0.25)
                                            time.sleep(0.3) 

                                            keyboard.press_and_release('4')
                                            time.sleep(0.3)
                                            

                                            pyautogui.click(x_registro, y_registro, button='left', clicks=1, interval=0.25)
                                            time.sleep(0.3) 

                                            keyboard.press_and_release('space')
                                            keyboard.press_and_release('space')
                                            keyboard.press_and_release('space')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('tab')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('n')
                                            time.sleep(0.3)
                                            

                                            pyautogui.click( x_registado_sim, y_registado_sim, button='left', clicks=1, interval=0.25)
                                            time.sleep(0.3) 

                                            keyboard.press_and_release('space')
                                            keyboard.press_and_release('space')
                                            keyboard.press_and_release('space')
                                            time.sleep(0.3)
                                            
                                            pyautogui.click(x_anterior, y_anterior, button='left', clicks=3, interval=0.25)
                                            time.sleep(0.1)
                                    


                                # Caso o estado seja "em processamento", retorna ao fluxo principal
                                elif 'em processament' in linha.lower():
                                        # print(f"Linha encontrada (Em Processamento): {linha}")
                                        encontrou_linha = True
                                        contador_em_processamento += 1

                                        match = re.match(r'^\d+', linha)
                                        if match:
                                            ot = match.group()
                                            print(f"Número da O.T.: {ot}")

                                            # Calcular a posição Y para o clique
                                            altura_linha = 20  # Ajuste conforme necessário para a altura da linha real
                                            y_pos = area[1] + (i * altura_linha) + 5
                                            x_pos = area[0] + 30  # Ajuste conforme necessário para a primeira coluna

                                            # Mover o mouse para a posição da linha específica e clicar
                                            pyautogui.moveTo(x_pos, y_pos)
                                            pyautogui.click()

                                            pyautogui.click(x_seguinte, y_seguinte, button='left', clicks=2, interval=0.25)
                                            time.sleep(1)

                                            pyautogui.click(x_registado, y_registado, button='left', clicks=2, interval=0.25)
                                            time.sleep(1)

                                            pyautogui.click(x_processamento_seguinte, y_processamento_seguinte, button='left', clicks=1, interval=0.25)
                                            time.sleep(5)

                                            pyautogui.click(x_registado_cmp1, y_registado_cmp1, button='left', clicks=1, interval=0.25)
                                            time.sleep(4)

                                            keyboard.press_and_release('r')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('tab')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('v')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('tab')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('2')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('tab')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('2')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('shift+tab')
                                            time.sleep(2)

                                            screenshot = pyautogui.screenshot(region=area2)
                                            texto_extraido = pytesseract.image_to_string(screenshot)
                                            
                                            texto_extraido = texto_extraido.strip()  # Remove espaços no início e no fim
                                            texto_extraido = texto_extraido.replace("\n", "")  # Remove quebras de linha

                                            # Mostrar o texto extraído para depuração
                                            print("Texto extraído:\n", texto_extraido)

                                            # Comparar o texto
                                            if texto_extraido == "22000 V":
                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)
                                                keyboard.press_and_release('2')
                                                time.sleep(0.3)
                                                
                                            else:
                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)
                                                
                                            keyboard.press_and_release('tab')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('s')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('tab')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('s')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('tab')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('r')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('tab')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('2')
                                            keyboard.press_and_release('0')
                                            keyboard.press_and_release('0')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('tab')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('4')
                                            keyboard.press_and_release('3')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('tab')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('4')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('tab')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('2')
                                            keyboard.press_and_release('0')
                                            keyboard.press_and_release('0')
                                            time.sleep(1)
                                            
                                            pyautogui.click(x_registado_cmp_potencia, y_registado_cmp1_potencia, button='left', clicks=1, interval=0.25)
                                            time.sleep(0.3)

                                            keyboard.press_and_release('2')
                                            keyboard.press_and_release('.')
                                            keyboard.press_and_release('2')
                                            time.sleep(0.3)
                                            
                                            pyautogui.click(x_propriedade, y_propriedade, button='left', clicks=2, interval=0.25)
                                            time.sleep(0.3)
                                            
                                            keyboard.press_and_release('backspace')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('1')
                                            keyboard.press_and_release('7')
                                            keyboard.press_and_release('2')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('tab')
                                            time.sleep(0.3)

                                            pyautogui.click(x_gis_x, y_gis_x, button='left', clicks=2, interval=0.25)
                                            time.sleep(0.3)
                                            
                                            keyboard.press_and_release('backspace')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('2')
                                            time.sleep(0.3)

                                            pyautogui.click(x_gis_y, y_gis_y, button='left', clicks=2, interval=0.25)
                                            time.sleep(0.3)
                                            
                                            keyboard.press_and_release('backspace')
                                            time.sleep(0.3)
                                            
                                            keyboard.press_and_release('1')
                                            time.sleep(0.3)

                                            pyautogui.click(x_estado_instalacao, y_estado_instalacao, button='left', clicks=1, interval=0.25)
                                            time.sleep(0.3) 

                                            keyboard.press_and_release('s')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('tab')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('s')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('tab')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('s')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('tab')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('s')
                                            time.sleep(0.3)

                                            pyautogui.click(x_num_luz, y_num_luz, button='left', clicks=1, interval=0.25)
                                            time.sleep(0.3) 

                                            keyboard.press_and_release('8')
                                            time.sleep(1)

                                            pyautogui.click(x_quartos, y_quartos, button='left', clicks=1, interval=0.25)
                                            time.sleep(0.3) 

                                            keyboard.press_and_release('4')
                                            time.sleep(0.3)

                                            pyautogui.click(x_registro, y_registro, button='left', clicks=1, interval=0.25)
                                            time.sleep(0.3) 

                                            keyboard.press_and_release('space')
                                            keyboard.press_and_release('space')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('tab')
                                            time.sleep(0.3)

                                            keyboard.press_and_release('n')
                                            time.sleep(0.3)

                                            pyautogui.click(x_registado_sim, y_registado_sim, button='left', clicks=1, interval=0.25)
                                            time.sleep(0.3) 

                                            keyboard.press_and_release('space')
                                            keyboard.press_and_release('space')
                                            time.sleep(0.3)

                                            pyautogui.click(x_anterior, y_anterior, button='left', clicks=3, interval=0.25)
                                            time.sleep(0.1)
                                

                                        # print("Retornando ao código principal.")
                                        break
                                    

                                if not encontrou_linha:
                                    contador_introduzido += 1
                                # print("Nenhuma linha correspondente encontrada. Pressionando Ctrl + I.")
                                    keyboard.press_and_release('ctrl+i')
                                    time.sleep(2)
                                    keyboard.press_and_release('tab')
                                    for _ in range(3):
                                        keyboard.press_and_release('i')
                                        time.sleep(1)
                                    for _ in range(3):    
                                        keyboard.press_and_release('tab')
                                    keyboard.write('rede perda')
                                    
                                    pyautogui.click(x_seguinte, y_seguinte, button='left', clicks=2, interval=0.25)
                                    time.sleep(1)
                                    
                                    for _ in range(4):    
                                        keyboard.press_and_release('space')
                                    time.sleep(1)
                                    
                                    pyautogui.click(x_processamento,y_processamento, button='left', clicks=2, interval=0.25)
                                    time.sleep(1)

                                    pyautogui.click(x_processamento_seguinte, y_processamento_seguinte, button='left', clicks=2, interval=0.25)
                                    time.sleep(4)

                                    pyautogui.click(x_imprimir, y_imprimir, button='left', clicks=1, interval=0.25)
                                    time.sleep(2)


                                    pyautogui.click( x_processamento_sim, y_processamento_sim, button='left', clicks=1, interval=0.25)
                                    time.sleep(4)

                                    pyautogui.click(x_cancelar, y_cancelar, button='left', clicks=1, interval=0.25)
                                    time.sleep(2)

                                    pyautogui.click(x_registado, y_registado, button='left', clicks=2, interval=0.25)
                                    time.sleep(1)

                                    pyautogui.click(x_registado_seguinte, y_registado_seguinte, button='left', clicks=2, interval=0.25)
                                    time.sleep(3)

                                    keyboard.press_and_release('space')

                                    pyautogui.click(x_registado_cmp1, y_registado_cmp1 , button='left', clicks=1, interval=0.25)
                                    time.sleep(5)

                                    keyboard.press_and_release('r')
                                    time.sleep(0.5)

                                    keyboard.press_and_release('tab')
                                    time.sleep(0.5)

                                    keyboard.press_and_release('v')
                                    time.sleep(0.5)

                                    keyboard.press_and_release('tab')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('2')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('tab')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('2')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('shift+tab')
                                    time.sleep(2)

                                    screenshot = pyautogui.screenshot(region=area2)
                                    texto_extraido = pytesseract.image_to_string(screenshot)
                                    
                                    texto_extraido = texto_extraido.strip()  # Remove espaços no início e no fim
                                    texto_extraido = texto_extraido.replace("\n", "")  # Remove quebras de linha

                                    # Mostrar o texto extraído para depuração
                                    print("Texto extraído:\n", texto_extraido)

                                    # Comparar o texto
                                    if texto_extraido == "22000 V":
                                        keyboard.press_and_release('tab')
                                        time.sleep(0.3)
                                        keyboard.press_and_release('2')
                                        time.sleep(0.3)
                                        
                                    else:
                                        keyboard.press_and_release('tab')
                                        time.sleep(0.3)
                                        
                                    keyboard.press_and_release('tab')
                                    time.sleep(0.3)
                                    keyboard.press_and_release('s')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('tab')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('s')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('tab')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('r')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('tab')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('2')
                                    keyboard.press_and_release('0')
                                    keyboard.press_and_release('0')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('tab')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('4')
                                    keyboard.press_and_release('3')
                                    time.sleep(0.3)
                
                                    keyboard.press_and_release('tab')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('4')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('tab')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('2')
                                    keyboard.press_and_release('0')
                                    keyboard.press_and_release('0')
                                    time.sleep(1)
                                    
                                    pyautogui.click(x_registado_cmp_potencia, y_registado_cmp1_potencia, button='left', clicks=1, interval=0.25)
                                    time.sleep(0.3)

                                    keyboard.press_and_release('2')
                                    keyboard.press_and_release('.')
                                    keyboard.press_and_release('2')
                                    time.sleep(0.3)
                                    
                                    pyautogui.click(x_propriedade, y_propriedade, button='left', clicks=2, interval=0.25)
                                    time.sleep(0.3)
                                    
                                    keyboard.press_and_release('backspace')
                                    time.sleep(0.3) 

                                    keyboard.press_and_release('1')
                                    keyboard.press_and_release('7')
                                    keyboard.press_and_release('2')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('tab')
                                    time.sleep(0.3)

                                    pyautogui.click(x_gis_x, y_gis_x, button='left', clicks=2, interval=0.25)
                                    time.sleep(0.3)
                                    
                                    keyboard.press_and_release('backspace')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('2')
                                    time.sleep(0.3)

                                    pyautogui.click(x_gis_y, y_gis_y, button='left', clicks=2, interval=0.25)
                                    time.sleep(0.3)
                                    
                                    keyboard.press_and_release('backspace')
                                    time.sleep(0.3)
                                    
                                    keyboard.press_and_release('1')
                                    time.sleep(0.3)

                                    
                                    pyautogui.click(x_estado_instalacao, y_estado_instalacao, button='left', clicks=1, interval=0.25)
                                    time.sleep(0.3) 

                                    keyboard.press_and_release('s')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('tab')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('s')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('tab')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('s')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('tab')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('s')
                                    time.sleep(0.3)
                                
                                    pyautogui.click(x_num_luz, y_num_luz , button='left', clicks=1, interval=0.25)
                                    time.sleep(0.3) 

                                    keyboard.press_and_release('8')
                                    time.sleep(1)

                                    pyautogui.click(x_quartos, y_quartos, button='left', clicks=1, interval=0.25)
                                    time.sleep(0.3) 

                                    keyboard.press_and_release('4')
                                    time.sleep(0.3)
                                    

                                    pyautogui.click(x_registro, y_registro, button='left', clicks=1, interval=0.25)
                                    time.sleep(0.3) 

                                    keyboard.press_and_release('space')
                                    keyboard.press_and_release('space')
                                    keyboard.press_and_release('space')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('tab')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('n')
                                    time.sleep(0.3)
                                    

                                    pyautogui.click( x_registado_sim, y_registado_sim, button='left', clicks=1, interval=0.25)
                                    time.sleep(0.3) 

                                    keyboard.press_and_release('space')
                                    keyboard.press_and_release('space')
                                    keyboard.press_and_release('space')
                                    time.sleep(0.3)
                                    
                                    pyautogui.click(x_anterior, y_anterior, button='left', clicks=3, interval=0.25)
                                    time.sleep(0.1)  







                            if len(linhas) % 3 == 2 or len(linhas) % 3 == 1:
                                print("\nLinhas (linha(s) detectada(s)):")
                                for linha in linhas:
                                    print(linha)  
                                
                                encontrou_linha = False

                                for i, linha in enumerate(linhas):
                                    if 'suspeita de fraude' in linha.lower():
                                        if 'registad' in linha.lower():
                                            print(f"Linha encontrada (Registado): {linha}")
                                            encontrou_linha = True
                                            contador_registados +=1

                                            # Extrair o número da O.T. usando regex (números no início da linha)
                                            match = re.match(r'\d+', linha)
                                            if match:
                                                ot = match.group()
                                                print(f"Número da O.T.: {ot}")

                                                # Calcular a posição Y para o clique
                                                altura_linha = 20  # Ajuste conforme necessário para a altura da linha real
                                                y_pos = area[1] + (i * altura_linha) + 5
                                                x_pos = area[0] + 30  # Ajuste conforme necessário para a primeira coluna

                                                # Mover o mouse para a posição da linha específica e clicar
                                                pyautogui.moveTo(x_pos, y_pos)
                                                pyautogui.click()

                                                print(f"Mouse movido e clique realizado na linha com O.T.: {ot}")

                                                pyautogui.click(x_seguinte, y_seguinte, button='left', clicks=2, interval=0.25)
                                                time.sleep(1)

                                                pyautogui.click(x_processamento,y_processamento, button='left', clicks=2, interval=0.25)
                                                time.sleep(1)

                                                pyautogui.click(x_processamento_seguinte, y_processamento_seguinte, button='left', clicks=2, interval=0.25)
                                                time.sleep(4)

                                                pyautogui.click(x_imprimir, y_imprimir, button='left', clicks=1, interval=0.25)
                                                time.sleep(2)


                                                pyautogui.click( x_processamento_sim, y_processamento_sim, button='left', clicks=1, interval=0.25)
                                                time.sleep(4)
                                                pyautogui.click(x_cancelar, y_cancelar, button='left', clicks=1, interval=0.25)
                                                time.sleep(2)

                                                pyautogui.click(x_registado, y_registado, button='left', clicks=2, interval=0.25)
                                                time.sleep(1)

                                                pyautogui.click(x_registado_seguinte, y_registado_seguinte, button='left', clicks=2, interval=0.25)
                                                time.sleep(3)

                                                keyboard.press_and_release('space')

                                                pyautogui.click(x_registado_cmp1, y_registado_cmp1 , button='left', clicks=1, interval=0.25)
                                                time.sleep(5)

                                                keyboard.press_and_release('r')
                                                time.sleep(0.5)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.5)

                                                keyboard.press_and_release('v')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('2')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('2')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('shift+tab')
                                                time.sleep(2)

                                                screenshot = pyautogui.screenshot(region=area2)
                                                texto_extraido = pytesseract.image_to_string(screenshot)
                                                
                                                texto_extraido = texto_extraido.strip()  # Remove espaços no início e no fim
                                                texto_extraido = texto_extraido.replace("\n", "")  # Remove quebras de linha

                                                # Mostrar o texto extraído para depuração
                                                print("Texto extraído:\n", texto_extraido)

                                                # Comparar o texto
                                                if texto_extraido == "22000 V":
                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)
                                                    keyboard.press_and_release('2')
                                                    time.sleep(0.3)
                                                    
                                                else:
                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)
                                                    
                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)
                                                keyboard.press_and_release('s')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('s')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('r')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('2')
                                                keyboard.press_and_release('0')
                                                keyboard.press_and_release('0')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('4')
                                                keyboard.press_and_release('3')
                                                time.sleep(0.3)
                            
                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('4')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('2')
                                                keyboard.press_and_release('0')
                                                keyboard.press_and_release('0')
                                                time.sleep(1)
                                                
                                                pyautogui.click(x_registado_cmp_potencia, y_registado_cmp1_potencia, button='left', clicks=1, interval=0.25)
                                                time.sleep(0.3)

                                                keyboard.press_and_release('2')
                                                keyboard.press_and_release('.')
                                                keyboard.press_and_release('2')
                                                time.sleep(0.3)
                                                
                                                pyautogui.click(x_propriedade, y_propriedade, button='left', clicks=2, interval=0.25)
                                                time.sleep(0.3)
                                                
                                                keyboard.press_and_release('backspace')
                                                time.sleep(0.3) 

                                                keyboard.press_and_release('1')
                                                keyboard.press_and_release('7')
                                                keyboard.press_and_release('2')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                pyautogui.click(x_gis_x, y_gis_x, button='left', clicks=2, interval=0.25)
                                                time.sleep(0.3)
                                                
                                                keyboard.press_and_release('backspace')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('2')
                                                time.sleep(0.3)

                                                pyautogui.click(x_gis_y, y_gis_y, button='left', clicks=2, interval=0.25)
                                                time.sleep(0.3)
                                                
                                                keyboard.press_and_release('backspace')
                                                time.sleep(0.3)
                                                
                                                keyboard.press_and_release('1')
                                                time.sleep(0.3)

                                                
                                                pyautogui.click(x_estado_instalacao, y_estado_instalacao, button='left', clicks=1, interval=0.25)
                                                time.sleep(0.3) 

                                                keyboard.press_and_release('s')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('s')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('s')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('s')
                                                time.sleep(0.3)
                                            
                                                pyautogui.click(x_num_luz, y_num_luz , button='left', clicks=1, interval=0.25)
                                                time.sleep(0.3) 

                                                keyboard.press_and_release('8')
                                                time.sleep(1)

                                                pyautogui.click(x_quartos, y_quartos, button='left', clicks=1, interval=0.25)
                                                time.sleep(0.3) 

                                                keyboard.press_and_release('4')
                                                time.sleep(0.3)
                                                

                                                pyautogui.click(x_registro, y_registro, button='left', clicks=1, interval=0.25)
                                                time.sleep(0.3) 

                                                keyboard.press_and_release('space')
                                                keyboard.press_and_release('space')
                                                keyboard.press_and_release('space')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('n')
                                                time.sleep(0.3)
                                                

                                                pyautogui.click( x_registado_sim, y_registado_sim, button='left', clicks=1, interval=0.25)
                                                time.sleep(0.3) 

                                                keyboard.press_and_release('space')
                                                keyboard.press_and_release('space')
                                                keyboard.press_and_release('space')
                                                time.sleep(0.3)
                                                
                                                pyautogui.click(x_anterior, y_anterior, button='left', clicks=3, interval=0.25)
                                                time.sleep(0.1)
                                    


                                        # Caso o estado seja "em processamento", retorna ao fluxo principal
                                        elif 'em processament' in linha.lower():
                                                # print(f"Linha encontrada (Em Processamento): {linha}")
                                                encontrou_linha = True
                                                contador_em_processamento += 1

                                                match = re.match(r'^\d+', linha)
                                                if match:
                                                    ot = match.group()
                                                    print(f"Número da O.T.: {ot}")

                                                    # Calcular a posição Y para o clique
                                                    altura_linha = 20  # Ajuste conforme necessário para a altura da linha real
                                                    y_pos = area[1] + (i * altura_linha) + 5
                                                    x_pos = area[0] + 30  # Ajuste conforme necessário para a primeira coluna

                                                    # Mover o mouse para a posição da linha específica e clicar
                                                    pyautogui.moveTo(x_pos, y_pos)
                                                    pyautogui.click()

                                                    pyautogui.click(x_seguinte, y_seguinte, button='left', clicks=2, interval=0.25)
                                                    time.sleep(1)

                                                    pyautogui.click(x_registado, y_registado, button='left', clicks=2, interval=0.25)
                                                    time.sleep(1)

                                                    pyautogui.click(x_processamento_seguinte, y_processamento_seguinte, button='left', clicks=1, interval=0.25)
                                                    time.sleep(5)

                                                    pyautogui.click(x_registado_cmp1, y_registado_cmp1, button='left', clicks=1, interval=0.25)
                                                    time.sleep(4)

                                                    keyboard.press_and_release('r')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('v')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('2')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('2')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('shift+tab')
                                                    time.sleep(2)

                                                    screenshot = pyautogui.screenshot(region=area2)
                                                    texto_extraido = pytesseract.image_to_string(screenshot)
                                                    
                                                    texto_extraido = texto_extraido.strip()  # Remove espaços no início e no fim
                                                    texto_extraido = texto_extraido.replace("\n", "")  # Remove quebras de linha

                                                    # Mostrar o texto extraído para depuração
                                                    print("Texto extraído:\n", texto_extraido)

                                                    # Comparar o texto
                                                    if texto_extraido == "22000 V":
                                                        keyboard.press_and_release('tab')
                                                        time.sleep(0.3)
                                                        keyboard.press_and_release('2')
                                                        time.sleep(0.3)
                                                        
                                                    else:
                                                        keyboard.press_and_release('tab')
                                                        time.sleep(0.3)
                                                        
                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('s')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('s')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('r')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('2')
                                                    keyboard.press_and_release('0')
                                                    keyboard.press_and_release('0')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('4')
                                                    keyboard.press_and_release('3')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('4')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('2')
                                                    keyboard.press_and_release('0')
                                                    keyboard.press_and_release('0')
                                                    time.sleep(1)
                                                    
                                                    pyautogui.click(x_registado_cmp_potencia, y_registado_cmp1_potencia, button='left', clicks=1, interval=0.25)
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('2')
                                                    keyboard.press_and_release('.')
                                                    keyboard.press_and_release('2')
                                                    time.sleep(0.3)
                                                    
                                                    pyautogui.click(x_propriedade, y_propriedade, button='left', clicks=2, interval=0.25)
                                                    time.sleep(0.3)
                                                    
                                                    keyboard.press_and_release('backspace')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('1')
                                                    keyboard.press_and_release('7')
                                                    keyboard.press_and_release('2')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)

                                                    pyautogui.click(x_gis_x, y_gis_x, button='left', clicks=2, interval=0.25)
                                                    time.sleep(0.3)
                                                    
                                                    keyboard.press_and_release('backspace')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('2')
                                                    time.sleep(0.3)

                                                    pyautogui.click(x_gis_y, y_gis_y, button='left', clicks=2, interval=0.25)
                                                    time.sleep(0.3)
                                                    
                                                    keyboard.press_and_release('backspace')
                                                    time.sleep(0.3)
                                                    
                                                    keyboard.press_and_release('1')
                                                    time.sleep(0.3)

                                                    pyautogui.click(x_estado_instalacao, y_estado_instalacao, button='left', clicks=1, interval=0.25)
                                                    time.sleep(0.3) 

                                                    keyboard.press_and_release('s')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('s')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('s')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('s')
                                                    time.sleep(0.3)

                                                    pyautogui.click(x_num_luz, y_num_luz, button='left', clicks=1, interval=0.25)
                                                    time.sleep(0.3) 

                                                    keyboard.press_and_release('8')
                                                    time.sleep(1)

                                                    pyautogui.click(x_quartos, y_quartos, button='left', clicks=1, interval=0.25)
                                                    time.sleep(0.3) 

                                                    keyboard.press_and_release('4')
                                                    time.sleep(0.3)

                                                    pyautogui.click(x_registro, y_registro, button='left', clicks=1, interval=0.25)
                                                    time.sleep(0.3) 

                                                    keyboard.press_and_release('space')
                                                    keyboard.press_and_release('space')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('n')
                                                    time.sleep(0.3)

                                                    pyautogui.click(x_registado_sim, y_registado_sim, button='left', clicks=1, interval=0.25)
                                                    time.sleep(0.3) 

                                                    keyboard.press_and_release('space')
                                                    keyboard.press_and_release('space')
                                                    time.sleep(0.3)

                                                    pyautogui.click(x_anterior, y_anterior, button='left', clicks=3, interval=0.25)
                                                    time.sleep(0.1)
                                        

                                                # print("Retornando ao código principal.")
                                                break
                                
                                if not encontrou_linha:
                                    contador_introduzido += 1
                                # print("Nenhuma linha correspondente encontrada. Pressionando Ctrl + I.")
                                    keyboard.press_and_release('ctrl+i')
                                    time.sleep(2)
                                    keyboard.press_and_release('tab')
                                    for _ in range(3):
                                        keyboard.press_and_release('i')
                                        time.sleep(1)
                                    for _ in range(3):    
                                        keyboard.press_and_release('tab')
                                    keyboard.write('rede perda')
                                    
                                    pyautogui.click(x_seguinte, y_seguinte, button='left', clicks=2, interval=0.25)
                                    time.sleep(1)
                                    
                                    for _ in range(4):    
                                        keyboard.press_and_release('space')
                                    time.sleep(1)
                                    
                                    pyautogui.click(x_processamento,y_processamento, button='left', clicks=2, interval=0.25)
                                    time.sleep(1)

                                    pyautogui.click(x_processamento_seguinte, y_processamento_seguinte, button='left', clicks=2, interval=0.25)
                                    time.sleep(4)

                                    pyautogui.click(x_imprimir, y_imprimir, button='left', clicks=1, interval=0.25)
                                    time.sleep(2)


                                    pyautogui.click( x_processamento_sim, y_processamento_sim, button='left', clicks=1, interval=0.25)
                                    time.sleep(4)

                                    pyautogui.click(x_cancelar, y_cancelar, button='left', clicks=1, interval=0.25)
                                    time.sleep(2)

                                    pyautogui.click(x_registado, y_registado, button='left', clicks=2, interval=0.25)
                                    time.sleep(1)

                                    pyautogui.click(x_registado_seguinte, y_registado_seguinte, button='left', clicks=2, interval=0.25)
                                    time.sleep(3)

                                    keyboard.press_and_release('space')

                                    pyautogui.click(x_registado_cmp1, y_registado_cmp1 , button='left', clicks=1, interval=0.25)
                                    time.sleep(5)

                                    keyboard.press_and_release('r')
                                    time.sleep(0.5)

                                    keyboard.press_and_release('tab')
                                    time.sleep(0.5)

                                    keyboard.press_and_release('v')
                                    time.sleep(0.5)

                                    keyboard.press_and_release('tab')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('2')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('tab')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('2')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('shift+tab')
                                    time.sleep(2)

                                    screenshot = pyautogui.screenshot(region=area2)
                                    texto_extraido = pytesseract.image_to_string(screenshot)
                                    
                                    texto_extraido = texto_extraido.strip()  # Remove espaços no início e no fim
                                    texto_extraido = texto_extraido.replace("\n", "")  # Remove quebras de linha

                                    # Mostrar o texto extraído para depuração
                                    print("Texto extraído:\n", texto_extraido)

                                    # Comparar o texto
                                    if texto_extraido == "22000 V":
                                        keyboard.press_and_release('tab')
                                        time.sleep(0.3)
                                        keyboard.press_and_release('2')
                                        time.sleep(0.3)
                                        
                                    else:
                                        keyboard.press_and_release('tab')
                                        time.sleep(0.3)
                                        
                                    keyboard.press_and_release('tab')
                                    time.sleep(0.3)
                                    keyboard.press_and_release('s')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('tab')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('s')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('tab')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('r')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('tab')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('2')
                                    keyboard.press_and_release('0')
                                    keyboard.press_and_release('0')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('tab')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('4')
                                    keyboard.press_and_release('3')
                                    time.sleep(0.3)
                
                                    keyboard.press_and_release('tab')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('4')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('tab')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('2')
                                    keyboard.press_and_release('0')
                                    keyboard.press_and_release('0')
                                    time.sleep(1)
                                    
                                    pyautogui.click(x_registado_cmp_potencia, y_registado_cmp1_potencia, button='left', clicks=1, interval=0.25)
                                    time.sleep(0.3)

                                    keyboard.press_and_release('2')
                                    keyboard.press_and_release('.')
                                    keyboard.press_and_release('2')
                                    time.sleep(0.3)
                                    
                                    pyautogui.click(x_propriedade, y_propriedade, button='left', clicks=2, interval=0.25)
                                    time.sleep(0.3)
                                    
                                    keyboard.press_and_release('backspace')
                                    time.sleep(0.3) 

                                    keyboard.press_and_release('1')
                                    keyboard.press_and_release('7')
                                    keyboard.press_and_release('2')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('tab')
                                    time.sleep(0.3)

                                    pyautogui.click(x_gis_x, y_gis_x, button='left', clicks=2, interval=0.25)
                                    time.sleep(0.3)
                                    
                                    keyboard.press_and_release('backspace')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('2')
                                    time.sleep(0.3)

                                    pyautogui.click(x_gis_y, y_gis_y, button='left', clicks=2, interval=0.25)
                                    time.sleep(0.3)
                                    
                                    keyboard.press_and_release('backspace')
                                    time.sleep(0.3)
                                    
                                    keyboard.press_and_release('1')
                                    time.sleep(0.3)

                                    
                                    pyautogui.click(x_estado_instalacao, y_estado_instalacao, button='left', clicks=1, interval=0.25)
                                    time.sleep(0.3) 

                                    keyboard.press_and_release('s')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('tab')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('s')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('tab')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('s')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('tab')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('s')
                                    time.sleep(0.3)
                                
                                    pyautogui.click(x_num_luz, y_num_luz , button='left', clicks=1, interval=0.25)
                                    time.sleep(0.3) 

                                    keyboard.press_and_release('8')
                                    time.sleep(1)

                                    pyautogui.click(x_quartos, y_quartos, button='left', clicks=1, interval=0.25)
                                    time.sleep(0.3) 

                                    keyboard.press_and_release('4')
                                    time.sleep(0.3)
                                    

                                    pyautogui.click(x_registro, y_registro, button='left', clicks=1, interval=0.25)
                                    time.sleep(0.3) 

                                    keyboard.press_and_release('space')
                                    keyboard.press_and_release('space')
                                    keyboard.press_and_release('space')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('tab')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('n')
                                    time.sleep(0.3)
                                    

                                    pyautogui.click( x_registado_sim, y_registado_sim, button='left', clicks=1, interval=0.25)
                                    time.sleep(0.3) 

                                    keyboard.press_and_release('space')
                                    keyboard.press_and_release('space')
                                    keyboard.press_and_release('space')
                                    time.sleep(0.3)
                                    
                                    pyautogui.click(x_anterior, y_anterior, button='left', clicks=3, interval=0.25)
                                    time.sleep(0.1)  







                    else:
                        if len(linhas) % 3 == 0:
                            if len(linhas) <= 3:    
                    # Caso especial: se há exatamente 3 linhas, não combinar, apenas exibir
                                print("\nLinhas (3 linhas detectadas):")
                                for linha in linhas:
                                    print(linha)
                                
                                encontrou_linha = False

                                for i, linha in enumerate(linhas):
                                    if 'suspeita de fraude' in linha.lower():
                                        if 'registad' in linha.lower():
                                            print(f"Linha encontrada (Registado): {linha}")
                                            encontrou_linha = True
                                            contador_registados +=1

                                            # Extrair o número da O.T. usando regex (números no início da linha)
                                            match = re.match(r'\d+', linha)
                                            if match:
                                                ot = match.group()
                                                print(f"Número da O.T.: {ot}")

                                                # Calcular a posição Y para o clique
                                                altura_linha = 20  # Ajuste conforme necessário para a altura da linha real
                                                y_pos = area[1] + (i * altura_linha) + 5
                                                x_pos = area[0] + 30  # Ajuste conforme necessário para a primeira coluna

                                                # Mover o mouse para a posição da linha específica e clicar
                                                pyautogui.moveTo(x_pos, y_pos)
                                                pyautogui.click()

                                                pyautogui.click(x_seguinte, y_seguinte, button='left', clicks=2, interval=0.25)
                                                time.sleep(1)

                                                pyautogui.click(x_processamento,y_processamento, button='left', clicks=2, interval=0.25)
                                                time.sleep(1)

                                                pyautogui.click(x_processamento_seguinte, y_processamento_seguinte, button='left', clicks=2, interval=0.25)
                                                time.sleep(4)

                                                pyautogui.click(x_imprimir, y_imprimir, button='left', clicks=1, interval=0.25)
                                                time.sleep(2)


                                                pyautogui.click( x_processamento_sim, y_processamento_sim, button='left', clicks=1, interval=0.25)
                                                time.sleep(4)

                                                pyautogui.click(x_cancelar, y_cancelar, button='left', clicks=1, interval=0.25)
                                                time.sleep(2)

                                                pyautogui.click(x_registado, y_registado, button='left', clicks=2, interval=0.25)
                                                time.sleep(1)

                                                pyautogui.click(x_registado_seguinte, y_registado_seguinte, button='left', clicks=2, interval=0.25)
                                                time.sleep(3)

                                                keyboard.press_and_release('space')

                                                pyautogui.click(x_registado_cmp1, y_registado_cmp1 , button='left', clicks=1, interval=0.25)
                                                time.sleep(5)

                                                keyboard.press_and_release('r')
                                                time.sleep(0.5)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.5)

                                                keyboard.press_and_release('v')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('2')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('2')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('shift+tab')
                                                time.sleep(2)

                                                screenshot = pyautogui.screenshot(region=area2)
                                                texto_extraido = pytesseract.image_to_string(screenshot)
                                                
                                                texto_extraido = texto_extraido.strip()  # Remove espaços no início e no fim
                                                texto_extraido = texto_extraido.replace("\n", "")  # Remove quebras de linha

                                                # Mostrar o texto extraído para depuração
                                                print("Texto extraído:\n", texto_extraido)

                                                # Comparar o texto
                                                if texto_extraido == "22000 V":
                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)
                                                    keyboard.press_and_release('2')
                                                    time.sleep(0.3)
                                                    
                                                else:
                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)
                                                    
                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)
                                                keyboard.press_and_release('s')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('s')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('r')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('2')
                                                keyboard.press_and_release('0')
                                                keyboard.press_and_release('0')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('4')
                                                keyboard.press_and_release('3')
                                                time.sleep(0.3)
                            
                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('4')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('2')
                                                keyboard.press_and_release('0')
                                                keyboard.press_and_release('0')
                                                time.sleep(1)
                                                
                                                pyautogui.click(x_registado_cmp_potencia, y_registado_cmp1_potencia, button='left', clicks=1, interval=0.25)
                                                time.sleep(0.3)

                                                keyboard.press_and_release('2')
                                                keyboard.press_and_release('.')
                                                keyboard.press_and_release('2')
                                                time.sleep(0.3)
                                                
                                                pyautogui.click(x_propriedade, y_propriedade, button='left', clicks=2, interval=0.25)
                                                time.sleep(0.3)
                                                
                                                keyboard.press_and_release('backspace')
                                                time.sleep(0.3) 

                                                keyboard.press_and_release('1')
                                                keyboard.press_and_release('7')
                                                keyboard.press_and_release('2')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                pyautogui.click(x_gis_x, y_gis_x, button='left', clicks=2, interval=0.25)
                                                time.sleep(0.3)
                                                
                                                keyboard.press_and_release('backspace')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('2')
                                                time.sleep(0.3)

                                                pyautogui.click(x_gis_y, y_gis_y, button='left', clicks=2, interval=0.25)
                                                time.sleep(0.3)
                                                
                                                keyboard.press_and_release('backspace')
                                                time.sleep(0.3)
                                                
                                                keyboard.press_and_release('1')
                                                time.sleep(0.3)

                                                
                                                pyautogui.click(x_estado_instalacao, y_estado_instalacao, button='left', clicks=1, interval=0.25)
                                                time.sleep(0.3) 

                                                keyboard.press_and_release('s')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('s')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('s')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('s')
                                                time.sleep(0.3)
                                            
                                                pyautogui.click(x_num_luz, y_num_luz , button='left', clicks=1, interval=0.25)
                                                time.sleep(0.3) 

                                                keyboard.press_and_release('8')
                                                time.sleep(1)

                                                pyautogui.click(x_quartos, y_quartos, button='left', clicks=1, interval=0.25)
                                                time.sleep(0.3) 

                                                keyboard.press_and_release('4')
                                                time.sleep(0.3)
                                                

                                                pyautogui.click(x_registro, y_registro, button='left', clicks=1, interval=0.25)
                                                time.sleep(0.3) 

                                                keyboard.press_and_release('space')
                                                keyboard.press_and_release('space')
                                                keyboard.press_and_release('space')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('n')
                                                time.sleep(0.3)
                                                

                                                pyautogui.click( x_registado_sim, y_registado_sim, button='left', clicks=1, interval=0.25)
                                                time.sleep(0.3) 

                                                keyboard.press_and_release('space')
                                                keyboard.press_and_release('space')
                                                keyboard.press_and_release('space')
                                                time.sleep(0.3)
                                                
                                                pyautogui.click(x_anterior, y_anterior, button='left', clicks=3, interval=0.25)
                                                time.sleep(0.1)

                                                
                                        elif 'em processament' in linha.lower():
                                        # print(f"Linha encontrada (Em Processamento): {linha}")
                                            encontrou_linha = True
                                            contador_em_processamento += 1

                                            match = re.match(r'\d+', linha)
                                            if match:
                                                ot = match.group()
                                                print(f"Número da O.T.: {ot}")

                                                # Calcular a posição Y para o clique
                                                altura_linha = 20  # Ajuste conforme necessário para a altura da linha real
                                                y_pos = area[1] + (i * altura_linha) + 5
                                                x_pos = area[0] + 30  # Ajuste conforme necessário para a primeira coluna

                                                # Mover o mouse para a posição da linha específica e clicar
                                                pyautogui.moveTo(x_pos, y_pos)
                                                pyautogui.click()

                                                pyautogui.click(x_seguinte, y_seguinte, button='left', clicks=2, interval=0.25)
                                                time.sleep(1)

                                                pyautogui.click(x_registado, y_registado, button='left', clicks=2, interval=0.25)
                                                time.sleep(1)

                                                pyautogui.click(x_processamento_seguinte, y_processamento_seguinte, button='left', clicks=1, interval=0.25)
                                                time.sleep(5)

                                                pyautogui.click(x_registado_cmp1, y_registado_cmp1, button='left', clicks=1, interval=0.25)
                                                time.sleep(2)

                                                keyboard.press_and_release('r')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('v')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('2')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('2')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('shift+tab')
                                                time.sleep(2)

                                                screenshot = pyautogui.screenshot(region=area2)
                                                texto_extraido = pytesseract.image_to_string(screenshot)
                                                
                                                texto_extraido = texto_extraido.strip()  # Remove espaços no início e no fim
                                                texto_extraido = texto_extraido.replace("\n", "")  # Remove quebras de linha

                                                # Mostrar o texto extraído para depuração
                                                print("Texto extraído:\n", texto_extraido)

                                                # Comparar o texto
                                                if texto_extraido == "22000 V":
                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)
                                                    keyboard.press_and_release('2')
                                                    time.sleep(0.3)
                                                    
                                                else:
                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)
                                                    
                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('s')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('s')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('r')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('2')
                                                keyboard.press_and_release('0')
                                                keyboard.press_and_release('0')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('4')
                                                keyboard.press_and_release('3')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('4')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('2')
                                                keyboard.press_and_release('0')
                                                keyboard.press_and_release('0')
                                                time.sleep(1)
                                                
                                                pyautogui.click(x_registado_cmp_potencia, y_registado_cmp1_potencia, button='left', clicks=1, interval=0.25)
                                                time.sleep(0.3)

                                                keyboard.press_and_release('2')
                                                keyboard.press_and_release('.')
                                                keyboard.press_and_release('2')
                                                time.sleep(0.3)
                                                
                                                pyautogui.click(x_propriedade, y_propriedade, button='left', clicks=2, interval=0.25)
                                                time.sleep(0.3)
                                                
                                                keyboard.press_and_release('backspace')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('1')
                                                keyboard.press_and_release('7')
                                                keyboard.press_and_release('2')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                pyautogui.click(x_gis_x, y_gis_x, button='left', clicks=2, interval=0.25)
                                                time.sleep(0.3)
                                                
                                                keyboard.press_and_release('backspace')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('2')
                                                time.sleep(0.3)

                                                pyautogui.click(x_gis_y, y_gis_y, button='left', clicks=2, interval=0.25)
                                                time.sleep(0.3)
                                                
                                                keyboard.press_and_release('backspace')
                                                time.sleep(0.3)
                                                
                                                keyboard.press_and_release('1')
                                                time.sleep(0.3)

                                                pyautogui.click(x_estado_instalacao, y_estado_instalacao, button='left', clicks=1, interval=0.25)
                                                time.sleep(0.3) 

                                                keyboard.press_and_release('s')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('s')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('s')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('s')
                                                time.sleep(0.3)

                                                pyautogui.click(x_num_luz, y_num_luz, button='left', clicks=1, interval=0.25)
                                                time.sleep(0.3) 

                                                keyboard.press_and_release('8')
                                                time.sleep(1)

                                                pyautogui.click(x_quartos, y_quartos, button='left', clicks=1, interval=0.25)
                                                time.sleep(0.3) 

                                                keyboard.press_and_release('4')
                                                time.sleep(0.3)

                                                pyautogui.click(x_registro, y_registro, button='left', clicks=1, interval=0.25)
                                                time.sleep(0.3) 

                                                keyboard.press_and_release('space')
                                                keyboard.press_and_release('space')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('n')
                                                time.sleep(0.3)

                                                pyautogui.click(x_registado_sim, y_registado_sim, button='left', clicks=1, interval=0.25)
                                                time.sleep(0.3) 

                                                keyboard.press_and_release('space')
                                                keyboard.press_and_release('space')
                                                time.sleep(0.3)

                                                pyautogui.click(x_anterior, y_anterior, button='left', clicks=3, interval=0.25)
                                                time.sleep(0.1)
                            

                                if not encontrou_linha:
                                    contador_introduzido += 1
                                    print("introducao com 3 linhas apenas")
                                # print("Nenhuma linha correspondente encontrada. Pressionando Ctrl + I.")
                                    keyboard.press_and_release('ctrl+i')
                                    time.sleep(2)
                                    keyboard.press_and_release('tab')
                                    for _ in range(3):
                                        keyboard.press_and_release('i')
                                        time.sleep(1)
                                    for _ in range(3):    
                                        keyboard.press_and_release('tab')
                                    keyboard.write('rede perda')
                                    
                                    pyautogui.click(x_seguinte, y_seguinte, button='left', clicks=2, interval=0.25)
                                    time.sleep(1)
                                    
                                    for _ in range(4):    
                                        keyboard.press_and_release('space')
                                    time.sleep(1)

                                    pyautogui.click(x_processamento,y_processamento, button='left', clicks=2, interval=0.25)
                                    time.sleep(1)

                                    pyautogui.click(x_processamento_seguinte, y_processamento_seguinte, button='left', clicks=2, interval=0.25)
                                    time.sleep(4)

                                    pyautogui.click(x_imprimir, y_imprimir, button='left', clicks=1, interval=0.25)
                                    time.sleep(2)


                                    pyautogui.click( x_processamento_sim, y_processamento_sim, button='left', clicks=1, interval=0.25)
                                    time.sleep(4)
                                    pyautogui.click(x_cancelar, y_cancelar, button='left', clicks=1, interval=0.25)
                                    time.sleep(2)

                                    pyautogui.click(x_registado, y_registado, button='left', clicks=2, interval=0.25)
                                    time.sleep(1)

                                    pyautogui.click(x_registado_seguinte, y_registado_seguinte, button='left', clicks=2, interval=0.25)
                                    time.sleep(3)

                                    keyboard.press_and_release('space')

                                    pyautogui.click(x_registado_cmp1, y_registado_cmp1 , button='left', clicks=1, interval=0.25)
                                    time.sleep(5)

                                    keyboard.press_and_release('r')
                                    time.sleep(0.5)

                                    keyboard.press_and_release('tab')
                                    time.sleep(0.5)

                                    keyboard.press_and_release('v')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('tab')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('2')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('tab')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('2')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('shift+tab')
                                    time.sleep(2)

                                    screenshot = pyautogui.screenshot(region=area2)
                                    texto_extraido = pytesseract.image_to_string(screenshot)
                                    
                                    texto_extraido = texto_extraido.strip()  # Remove espaços no início e no fim
                                    texto_extraido = texto_extraido.replace("\n", "")  # Remove quebras de linha

                                    # Mostrar o texto extraído para depuração
                                    print("Texto extraído:\n", texto_extraido)

                                    # Comparar o texto
                                    if texto_extraido == "22000 V":
                                        keyboard.press_and_release('tab')
                                        time.sleep(0.3)
                                        keyboard.press_and_release('2')
                                        time.sleep(0.3)
                                        
                                    else:
                                        keyboard.press_and_release('tab')
                                        time.sleep(0.3)
                                        
                                    keyboard.press_and_release('tab')
                                    time.sleep(0.3)
                                    keyboard.press_and_release('s')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('tab')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('s')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('tab')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('r')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('tab')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('2')
                                    keyboard.press_and_release('0')
                                    keyboard.press_and_release('0')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('tab')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('4')
                                    keyboard.press_and_release('3')
                                    time.sleep(0.3)
                
                                    keyboard.press_and_release('tab')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('4')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('tab')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('2')
                                    keyboard.press_and_release('0')
                                    keyboard.press_and_release('0')
                                    time.sleep(1)
                                    
                                    pyautogui.click(x_registado_cmp_potencia, y_registado_cmp1_potencia, button='left', clicks=1, interval=0.25)
                                    time.sleep(0.3)

                                    keyboard.press_and_release('2')
                                    keyboard.press_and_release('.')
                                    keyboard.press_and_release('2')
                                    time.sleep(0.3)
                                    
                                    pyautogui.click(x_propriedade, y_propriedade, button='left', clicks=2, interval=0.25)
                                    time.sleep(0.3)
                                    
                                    keyboard.press_and_release('backspace')
                                    time.sleep(0.3) 

                                    keyboard.press_and_release('1')
                                    keyboard.press_and_release('7')
                                    keyboard.press_and_release('2')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('tab')
                                    time.sleep(0.3)

                                    pyautogui.click(x_gis_x, y_gis_x, button='left', clicks=2, interval=0.25)
                                    time.sleep(0.3)
                                    
                                    keyboard.press_and_release('backspace')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('2')
                                    time.sleep(0.3)

                                    pyautogui.click(x_gis_y, y_gis_y, button='left', clicks=2, interval=0.25)
                                    time.sleep(0.3)
                                    
                                    keyboard.press_and_release('backspace')
                                    time.sleep(0.3)
                                    
                                    keyboard.press_and_release('1')
                                    time.sleep(0.3)

                                    
                                    pyautogui.click(x_estado_instalacao, y_estado_instalacao, button='left', clicks=1, interval=0.25)
                                    time.sleep(0.3) 

                                    keyboard.press_and_release('s')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('tab')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('s')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('tab')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('s')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('tab')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('s')
                                    time.sleep(0.3)
                                
                                    pyautogui.click(x_num_luz, y_num_luz , button='left', clicks=1, interval=0.25)
                                    time.sleep(0.3) 

                                    keyboard.press_and_release('8')
                                    time.sleep(1)

                                    pyautogui.click(x_quartos, y_quartos, button='left', clicks=1, interval=0.25)
                                    time.sleep(0.3) 

                                    keyboard.press_and_release('4')
                                    time.sleep(0.3)
                                    

                                    pyautogui.click(x_registro, y_registro, button='left', clicks=1, interval=0.25)
                                    time.sleep(0.3) 

                                    keyboard.press_and_release('space')
                                    keyboard.press_and_release('space')
                                    keyboard.press_and_release('space')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('tab')
                                    time.sleep(0.3)

                                    keyboard.press_and_release('n')
                                    time.sleep(0.3)
                                    

                                    pyautogui.click( x_registado_sim, y_registado_sim, button='left', clicks=1, interval=0.25)
                                    time.sleep(0.3) 

                                    keyboard.press_and_release('space')
                                    keyboard.press_and_release('space')
                                    keyboard.press_and_release('space')
                                    time.sleep(0.3)
                                    
                                    pyautogui.click(x_anterior, y_anterior, button='left', clicks=3, interval=0.25)
                                    time.sleep(0.1)
                                    continue           
                            else: 
                                
                                total_grupos = len(linhas) // 3
                                tipos = linhas[:total_grupos]
                                estados = linhas[total_grupos:2 * total_grupos]
                                centros = linhas[2 * total_grupos:]

                                # Combinar os dados correspondentes
                                linhas_combinadas = []
                                for i in range(total_grupos):
                                    linha_completa = f"{tipos[i]} {estados[i]} {centros[i]}"
                                    linhas_combinadas.append(linha_completa)

                                # Mostrar as linhas combinadas
                                print("\nLinhas Combinadas:")
                                for linha in linhas_combinadas:
                                    print(linha)

                                # Definir os comportamentos baseados em palavras-chave
                                encontrou_linha = False

                                for i, linha in enumerate(linhas_combinadas):
                                    if 'suspeita de fraude' in linha.lower():
                                        if 'registad' in linha.lower():
                                            print(f"Linha encontrada (Registado): {linha}")
                                            encontrou_linha = True
                                            contador_registados +=1

                                            # Extrair o número da O.T. usando regex (números no início da linha)
                                            match = re.match(r'^\d+', linha)
                                            if match:
                                                ot = match.group()
                                                print(f"Número da O.T.: {ot}")

                                                # Calcular a posição Y para o clique
                                                altura_linha = 20  # Ajuste conforme necessário para a altura da linha real
                                                y_pos = area[1] + (i * altura_linha) + 5
                                                x_pos = area[0] + 30  # Ajuste conforme necessário para a primeira coluna

                                                # Mover o mouse para a posição da linha específica e clicar
                                                pyautogui.moveTo(x_pos, y_pos)
                                                pyautogui.click()
                                                print(f"Mouse movido e clique realizado na linha com O.T.: {ot}")

                                                pyautogui.click(x_seguinte, y_seguinte, button='left', clicks=2, interval=0.25)
                                                time.sleep(1)

                                                pyautogui.click(x_processamento,y_processamento, button='left', clicks=2, interval=0.25)
                                                time.sleep(1)

                                                pyautogui.click(x_processamento_seguinte, y_processamento_seguinte, button='left', clicks=2, interval=0.25)
                                                time.sleep(4)

                                                pyautogui.click(x_imprimir, y_imprimir, button='left', clicks=1, interval=0.25)
                                                time.sleep(2)


                                                pyautogui.click( x_processamento_sim, y_processamento_sim, button='left', clicks=1, interval=0.25)
                                                time.sleep(4)

                                                pyautogui.click(x_cancelar, y_cancelar, button='left', clicks=1, interval=0.25)
                                                time.sleep(2)

                                                pyautogui.click(x_registado, y_registado, button='left', clicks=2, interval=0.25)
                                                time.sleep(1)

                                                pyautogui.click(x_registado_seguinte, y_registado_seguinte, button='left', clicks=2, interval=0.25)
                                                time.sleep(3)

                                                keyboard.press_and_release('space')

                                                pyautogui.click(x_registado_cmp1, y_registado_cmp1 , button='left', clicks=1, interval=0.25)
                                                time.sleep(5)

                                                keyboard.press_and_release('r')
                                                time.sleep(0.5)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.5)

                                                keyboard.press_and_release('v')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('2')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('2')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('shift+tab')
                                                time.sleep(2)

                                                screenshot = pyautogui.screenshot(region=area2)
                                                texto_extraido = pytesseract.image_to_string(screenshot)
                                                
                                                texto_extraido = texto_extraido.strip()  # Remove espaços no início e no fim
                                                texto_extraido = texto_extraido.replace("\n", "")  # Remove quebras de linha

                                                # Mostrar o texto extraído para depuração
                                                print("Texto extraído:\n", texto_extraido)

                                                # Comparar o texto
                                                if texto_extraido == "22000 V":
                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)
                                                    keyboard.press_and_release('2')
                                                    time.sleep(0.3)
                                                    
                                                else:
                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)
                                                    
                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)
                                                keyboard.press_and_release('s')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('s')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('r')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('2')
                                                keyboard.press_and_release('0')
                                                keyboard.press_and_release('0')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('4')
                                                keyboard.press_and_release('3')
                                                time.sleep(0.3)
                            
                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('4')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('2')
                                                keyboard.press_and_release('0')
                                                keyboard.press_and_release('0')
                                                time.sleep(1)
                                                
                                                pyautogui.click(x_registado_cmp_potencia, y_registado_cmp1_potencia, button='left', clicks=1, interval=0.25)
                                                time.sleep(0.3)

                                                keyboard.press_and_release('2')
                                                keyboard.press_and_release('.')
                                                keyboard.press_and_release('2')
                                                time.sleep(0.3)
                                                
                                                pyautogui.click(x_propriedade, y_propriedade, button='left', clicks=2, interval=0.25)
                                                time.sleep(0.3)
                                                
                                                keyboard.press_and_release('backspace')
                                                time.sleep(0.3) 

                                                keyboard.press_and_release('1')
                                                keyboard.press_and_release('7')
                                                keyboard.press_and_release('2')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                pyautogui.click(x_gis_x, y_gis_x, button='left', clicks=2, interval=0.25)
                                                time.sleep(0.3)
                                                
                                                keyboard.press_and_release('backspace')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('2')
                                                time.sleep(0.3)

                                                pyautogui.click(x_gis_y, y_gis_y, button='left', clicks=2, interval=0.25)
                                                time.sleep(0.3)
                                                
                                                keyboard.press_and_release('backspace')
                                                time.sleep(0.3)
                                                
                                                keyboard.press_and_release('1')
                                                time.sleep(0.3)

                                                
                                                pyautogui.click(x_estado_instalacao, y_estado_instalacao, button='left', clicks=1, interval=0.25)
                                                time.sleep(0.3) 

                                                keyboard.press_and_release('s')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('s')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('s')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('s')
                                                time.sleep(0.3)
                                            
                                                pyautogui.click(x_num_luz, y_num_luz , button='left', clicks=1, interval=0.25)
                                                time.sleep(0.3) 

                                                keyboard.press_and_release('8')
                                                time.sleep(1)

                                                pyautogui.click(x_quartos, y_quartos, button='left', clicks=1, interval=0.25)
                                                time.sleep(0.3) 

                                                keyboard.press_and_release('4')
                                                time.sleep(0.3)
                                                

                                                pyautogui.click(x_registro, y_registro, button='left', clicks=1, interval=0.25)
                                                time.sleep(0.3) 

                                                keyboard.press_and_release('space')
                                                keyboard.press_and_release('space')
                                                keyboard.press_and_release('space')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('n')
                                                time.sleep(0.3)
                                                

                                                pyautogui.click( x_registado_sim, y_registado_sim, button='left', clicks=1, interval=0.25)
                                                time.sleep(0.3) 

                                                keyboard.press_and_release('space')
                                                keyboard.press_and_release('space')
                                                keyboard.press_and_release('space')
                                                time.sleep(0.3)
                                                
                                                pyautogui.click(x_anterior, y_anterior, button='left', clicks=3, interval=0.25)
                                                time.sleep(0.1)
                                        


                                    # Caso o estado seja "em processamento", retorna ao fluxo principal
                                        elif 'em processament' in linha.lower():
                                            # print(f"Linha encontrada (Em Processamento): {linha}")
                                            encontrou_linha = True
                                            contador_em_processamento += 1

                                            match = re.match(r'\d+', linha)
                                            if match:
                                                ot = match.group()
                                                print(f"Número da O.T.: {ot}")

                                                # Calcular a posição Y para o clique
                                                altura_linha = 20  # Ajuste conforme necessário para a altura da linha real
                                                y_pos = area[1] + (i * altura_linha) + 5
                                                x_pos = area[0] + 30  # Ajuste conforme necessário para a primeira coluna

                                                # Mover o mouse para a posição da linha específica e clicar
                                                pyautogui.moveTo(x_pos, y_pos)
                                                pyautogui.click()

                                                pyautogui.click(x_seguinte, y_seguinte, button='left', clicks=2, interval=0.25)
                                                time.sleep(1)

                                                pyautogui.click(x_registado, y_registado, button='left', clicks=2, interval=0.25)
                                                time.sleep(1)

                                                pyautogui.click(x_processamento_seguinte, y_processamento_seguinte, button='left', clicks=1, interval=0.25)
                                                time.sleep(5)

                                                pyautogui.click(x_registado_cmp1, y_registado_cmp1, button='left', clicks=1, interval=0.25)
                                                time.sleep(2)

                                                keyboard.press_and_release('r')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('v')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('2')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('2')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('shift+tab')
                                                time.sleep(2)

                                                screenshot = pyautogui.screenshot(region=area2)
                                                texto_extraido = pytesseract.image_to_string(screenshot)
                                                
                                                texto_extraido = texto_extraido.strip()  # Remove espaços no início e no fim
                                                texto_extraido = texto_extraido.replace("\n", "")  # Remove quebras de linha

                                                # Mostrar o texto extraído para depuração
                                                print("Texto extraído:\n", texto_extraido)

                                                # Comparar o texto
                                                if texto_extraido == "22000 V":
                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)
                                                    keyboard.press_and_release('2')
                                                    time.sleep(0.3)
                                                    
                                                else:
                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)
                                                    
                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('s')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('s')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('r')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('2')
                                                keyboard.press_and_release('0')
                                                keyboard.press_and_release('0')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('4')
                                                keyboard.press_and_release('3')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('4')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('2')
                                                keyboard.press_and_release('0')
                                                keyboard.press_and_release('0')
                                                time.sleep(1)
                                                
                                                pyautogui.click(x_registado_cmp_potencia, y_registado_cmp1_potencia, button='left', clicks=1, interval=0.25)
                                                time.sleep(0.3)

                                                keyboard.press_and_release('2')
                                                keyboard.press_and_release('.')
                                                keyboard.press_and_release('2')
                                                time.sleep(0.3)
                                                
                                                pyautogui.click(x_propriedade, y_propriedade, button='left', clicks=2, interval=0.25)
                                                time.sleep(0.3)
                                                
                                                keyboard.press_and_release('backspace')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('1')
                                                keyboard.press_and_release('7')
                                                keyboard.press_and_release('2')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                pyautogui.click(x_gis_x, y_gis_x, button='left', clicks=2, interval=0.25)
                                                time.sleep(0.3)
                                                
                                                keyboard.press_and_release('backspace')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('2')
                                                time.sleep(0.3)

                                                pyautogui.click(x_gis_y, y_gis_y, button='left', clicks=2, interval=0.25)
                                                time.sleep(0.3)
                                                
                                                keyboard.press_and_release('backspace')
                                                time.sleep(0.3)
                                                
                                                keyboard.press_and_release('1')
                                                time.sleep(0.3)

                                                pyautogui.click(x_estado_instalacao, y_estado_instalacao, button='left', clicks=1, interval=0.25)
                                                time.sleep(0.3) 

                                                keyboard.press_and_release('s')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('s')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('s')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('s')
                                                time.sleep(0.3)

                                                pyautogui.click(x_num_luz, y_num_luz, button='left', clicks=1, interval=0.25)
                                                time.sleep(0.3) 

                                                keyboard.press_and_release('8')
                                                time.sleep(1)

                                                pyautogui.click(x_quartos, y_quartos, button='left', clicks=1, interval=0.25)
                                                time.sleep(0.3) 

                                                keyboard.press_and_release('4')
                                                time.sleep(0.3)

                                                pyautogui.click(x_registro, y_registro, button='left', clicks=1, interval=0.25)
                                                time.sleep(0.3) 

                                                keyboard.press_and_release('space')
                                                keyboard.press_and_release('space')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('tab')
                                                time.sleep(0.3)

                                                keyboard.press_and_release('n')
                                                time.sleep(0.3)

                                                pyautogui.click(x_registado_sim, y_registado_sim, button='left', clicks=1, interval=0.25)
                                                time.sleep(0.3) 

                                                keyboard.press_and_release('space')
                                                keyboard.press_and_release('space')
                                                time.sleep(0.3)

                                                pyautogui.click(x_anterior, y_anterior, button='left', clicks=3, interval=0.25)
                                                time.sleep(0.1)
                            

                                            # print("Retornando ao código principal.")
                                                break

                            # Se nenhuma linha correspondente for encontrada, pressionar Ctrl + I
                                if not encontrou_linha:
                                    pyautogui.click(click_position)
                                    time.sleep(0.1)

                                    keyboard.press_and_release('page down')
                                    time.sleep(0.5)
                                    # print("Nenhuma linha correspondente encontrada. Pressionando Ctrl + I.")
                                    screenshot = pyautogui.screenshot(region=area)

                        # Usar o pytesseract para extrair texto da captura de tela
                                    texto_extraido = pytesseract.image_to_string(screenshot)

                                    # Mostrar o texto extraído (para depuração)
                                    print("Texto extraído:\n", texto_extraido)

                                    # Dividir o texto extraído em linhas, removendo espaços em branco extras
                                    linhas = [linha.strip() for linha in texto_extraido.split('\n') if linha.strip()]
                                

                                    total_grupos = len(linhas) // 3
                                    tipos = linhas[:total_grupos]
                                    estados = linhas[total_grupos:2 * total_grupos]
                                    centros = linhas[2 * total_grupos:]

                                    # Combinar os dados correspondentes
                                    linhas_combinadas = []
                                    for i in range(total_grupos):
                                        linha_completa = f"{tipos[i]} {estados[i]} {centros[i]}"
                                        linhas_combinadas.append(linha_completa)

                                    # Mostrar as linhas combinadas
                                    print("\nLinhas Combinadas:")
                                    for linha in linhas_combinadas:
                                        print(linha)

                                    # Definir os comportamentos baseados em palavras-chave
                                    encontrou_linha = False

                                    for i, linha in enumerate(linhas_combinadas):
                                        if 'suspeita de fraude' in linha.lower():
                                            if 'registad' in linha.lower():
                                                print(f"Linha encontrada (Registado): {linha}")
                                                encontrou_linha = True
                                                contador_registados +=1

                                                # Extrair o número da O.T. usando regex (números no início da linha)
                                                match = re.match(r'^\d+', linha)
                                                if match:
                                                    ot = match.group()
                                                    print(f"Número da O.T.: {ot}")

                                                    # Calcular a posição Y para o clique
                                                    altura_linha = 20  # Ajuste conforme necessário para a altura da linha real
                                                    y_pos = area[1] + (i * altura_linha) + 5
                                                    x_pos = area[0] + 30  # Ajuste conforme necessário para a primeira coluna

                                                    # Mover o mouse para a posição da linha específica e clicar
                                                    pyautogui.moveTo(x_pos, y_pos)
                                                    pyautogui.click()
                                                    print(f"Mouse movido e clique realizado na linha com O.T.: {ot}")

                                                    pyautogui.click(x_seguinte, y_seguinte, button='left', clicks=2, interval=0.25)
                                                    time.sleep(1)

                                                    pyautogui.click(x_processamento,y_processamento, button='left', clicks=2, interval=0.25)
                                                    time.sleep(1)

                                                    pyautogui.click(x_processamento_seguinte, y_processamento_seguinte, button='left', clicks=2, interval=0.25)
                                                    time.sleep(4)

                                                    pyautogui.click(x_imprimir, y_imprimir, button='left', clicks=1, interval=0.25)
                                                    time.sleep(2)


                                                    pyautogui.click( x_processamento_sim, y_processamento_sim, button='left', clicks=1, interval=0.25)
                                                    time.sleep(4)
                                                    pyautogui.click(x_cancelar, y_cancelar, button='left', clicks=1, interval=0.25)
                                                    time.sleep(2)

                                                    pyautogui.click(x_registado, y_registado, button='left', clicks=2, interval=0.25)
                                                    time.sleep(1)

                                                    pyautogui.click(x_registado_seguinte, y_registado_seguinte, button='left', clicks=2, interval=0.25)
                                                    time.sleep(3)

                                                    keyboard.press_and_release('space')

                                                    pyautogui.click(x_registado_cmp1, y_registado_cmp1 , button='left', clicks=1, interval=0.25)
                                                    time.sleep(5)

                                                    keyboard.press_and_release('r')
                                                    time.sleep(0.5)

                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.5)

                                                    keyboard.press_and_release('v')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('2')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('2')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('shift+tab')
                                                    time.sleep(2)

                                                    screenshot = pyautogui.screenshot(region=area2)
                                                    texto_extraido = pytesseract.image_to_string(screenshot)
                                                    
                                                    texto_extraido = texto_extraido.strip()  # Remove espaços no início e no fim
                                                    texto_extraido = texto_extraido.replace("\n", "")  # Remove quebras de linha

                                                    # Mostrar o texto extraído para depuração
                                                    print("Texto extraído:\n", texto_extraido)

                                                    # Comparar o texto
                                                    if texto_extraido == "22000 V":
                                                        keyboard.press_and_release('tab')
                                                        time.sleep(0.3)
                                                        keyboard.press_and_release('2')
                                                        time.sleep(0.3)
                                                        
                                                    else:
                                                        keyboard.press_and_release('tab')
                                                        time.sleep(0.3)
                                                        
                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)
                                                    keyboard.press_and_release('s')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('s')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('r')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('2')
                                                    keyboard.press_and_release('0')
                                                    keyboard.press_and_release('0')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('4')
                                                    keyboard.press_and_release('3')
                                                    time.sleep(0.3)
                                
                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('4')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('2')
                                                    keyboard.press_and_release('0')
                                                    keyboard.press_and_release('0')
                                                    time.sleep(1)
                                                    
                                                    pyautogui.click(x_registado_cmp_potencia, y_registado_cmp1_potencia, button='left', clicks=1, interval=0.25)
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('2')
                                                    keyboard.press_and_release('.')
                                                    keyboard.press_and_release('2')
                                                    time.sleep(0.3)
                                                    
                                                    pyautogui.click(x_propriedade, y_propriedade, button='left', clicks=2, interval=0.25)
                                                    time.sleep(0.3)
                                                    
                                                    keyboard.press_and_release('backspace')
                                                    time.sleep(0.3) 

                                                    keyboard.press_and_release('1')
                                                    keyboard.press_and_release('7')
                                                    keyboard.press_and_release('2')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)

                                                    pyautogui.click(x_gis_x, y_gis_x, button='left', clicks=2, interval=0.25)
                                                    time.sleep(0.3)
                                                    
                                                    keyboard.press_and_release('backspace')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('2')
                                                    time.sleep(0.3)

                                                    pyautogui.click(x_gis_y, y_gis_y, button='left', clicks=2, interval=0.25)
                                                    time.sleep(0.3)
                                                    
                                                    keyboard.press_and_release('backspace')
                                                    time.sleep(0.3)
                                                    
                                                    keyboard.press_and_release('1')
                                                    time.sleep(0.3)

                                                    
                                                    pyautogui.click(x_estado_instalacao, y_estado_instalacao, button='left', clicks=1, interval=0.25)
                                                    time.sleep(0.3) 

                                                    keyboard.press_and_release('s')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('s')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('s')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('s')
                                                    time.sleep(0.3)
                                                
                                                    pyautogui.click(x_num_luz, y_num_luz , button='left', clicks=1, interval=0.25)
                                                    time.sleep(0.3) 

                                                    keyboard.press_and_release('8')
                                                    time.sleep(1)

                                                    pyautogui.click(x_quartos, y_quartos, button='left', clicks=1, interval=0.25)
                                                    time.sleep(0.3) 

                                                    keyboard.press_and_release('4')
                                                    time.sleep(0.3)
                                                    

                                                    pyautogui.click(x_registro, y_registro, button='left', clicks=1, interval=0.25)
                                                    time.sleep(0.3) 

                                                    keyboard.press_and_release('space')
                                                    keyboard.press_and_release('space')
                                                    keyboard.press_and_release('space')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('n')
                                                    time.sleep(0.3)
                                                    

                                                    pyautogui.click( x_registado_sim, y_registado_sim, button='left', clicks=1, interval=0.25)
                                                    time.sleep(0.3) 

                                                    keyboard.press_and_release('space')
                                                    keyboard.press_and_release('space')
                                                    keyboard.press_and_release('space')
                                                    time.sleep(0.3)
                                                    
                                                    pyautogui.click(x_anterior, y_anterior, button='left', clicks=3, interval=0.25)
                                                    time.sleep(0.1)
                                            


                                        # Caso o estado seja "em processamento", retorna ao fluxo principal
                                            elif 'em processament' in linha.lower():
                                                # print(f"Linha encontrada (Em Processamento): {linha}")
                                                encontrou_linha = True
                                                contador_em_processamento += 1

                                                match = re.match(r'\d+', linha)
                                                if match:
                                                    ot = match.group()
                                                    print(f"Número da O.T.: {ot}")

                                                    # Calcular a posição Y para o clique
                                                    altura_linha = 20  # Ajuste conforme necessário para a altura da linha real
                                                    y_pos = area[1] + (i * altura_linha) + 5
                                                    x_pos = area[0] + 30  # Ajuste conforme necessário para a primeira coluna

                                                    # Mover o mouse para a posição da linha específica e clicar
                                                    pyautogui.moveTo(x_pos, y_pos)
                                                    pyautogui.click()

                                                    pyautogui.click(x_seguinte, y_seguinte, button='left', clicks=2, interval=0.25)
                                                    time.sleep(1)

                                                    pyautogui.click(x_registado, y_registado, button='left', clicks=2, interval=0.25)
                                                    time.sleep(1)

                                                    pyautogui.click(x_processamento_seguinte, y_processamento_seguinte, button='left', clicks=1, interval=0.25)
                                                    time.sleep(5)

                                                    pyautogui.click(x_registado_cmp1, y_registado_cmp1, button='left', clicks=1, interval=0.25)
                                                    time.sleep(2)

                                                    keyboard.press_and_release('r')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('v')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('2')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('2')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('shift+tab')
                                                    time.sleep(2)

                                                    screenshot = pyautogui.screenshot(region=area2)
                                                    texto_extraido = pytesseract.image_to_string(screenshot)
                                                    
                                                    texto_extraido = texto_extraido.strip()  # Remove espaços no início e no fim
                                                    texto_extraido = texto_extraido.replace("\n", "")  # Remove quebras de linha

                                                    # Mostrar o texto extraído para depuração
                                                    print("Texto extraído:\n", texto_extraido)

                                                    # Comparar o texto
                                                    if texto_extraido == "22000 V":
                                                        keyboard.press_and_release('tab')
                                                        time.sleep(0.3)
                                                        keyboard.press_and_release('2')
                                                        time.sleep(0.3)
                                                        
                                                    else:
                                                        keyboard.press_and_release('tab')
                                                        time.sleep(0.3)
                                                        
                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('s')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('s')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('r')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('2')
                                                    keyboard.press_and_release('0')
                                                    keyboard.press_and_release('0')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('4')
                                                    keyboard.press_and_release('3')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('4')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('2')
                                                    keyboard.press_and_release('0')
                                                    keyboard.press_and_release('0')
                                                    time.sleep(1)
                                                    
                                                    pyautogui.click(x_registado_cmp_potencia, y_registado_cmp1_potencia, button='left', clicks=1, interval=0.25)
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('2')
                                                    keyboard.press_and_release('.')
                                                    keyboard.press_and_release('2')
                                                    time.sleep(0.3)
                                                    
                                                    pyautogui.click(x_propriedade, y_propriedade, button='left', clicks=2, interval=0.25)
                                                    time.sleep(0.3)
                                                    
                                                    keyboard.press_and_release('backspace')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('1')
                                                    keyboard.press_and_release('7')
                                                    keyboard.press_and_release('2')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)

                                                    pyautogui.click(x_gis_x, y_gis_x, button='left', clicks=2, interval=0.25)
                                                    time.sleep(0.3)
                                                    
                                                    keyboard.press_and_release('backspace')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('2')
                                                    time.sleep(0.3)

                                                    pyautogui.click(x_gis_y, y_gis_y, button='left', clicks=2, interval=0.25)
                                                    time.sleep(0.3)
                                                    
                                                    keyboard.press_and_release('backspace')
                                                    time.sleep(0.3)
                                                    
                                                    keyboard.press_and_release('1')
                                                    time.sleep(0.3)

                                                    pyautogui.click(x_estado_instalacao, y_estado_instalacao, button='left', clicks=1, interval=0.25)
                                                    time.sleep(0.3) 

                                                    keyboard.press_and_release('s')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('s')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('s')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('s')
                                                    time.sleep(0.3)

                                                    pyautogui.click(x_num_luz, y_num_luz, button='left', clicks=1, interval=0.25)
                                                    time.sleep(0.3) 

                                                    keyboard.press_and_release('8')
                                                    time.sleep(1)

                                                    pyautogui.click(x_quartos, y_quartos, button='left', clicks=1, interval=0.25)
                                                    time.sleep(0.3) 

                                                    keyboard.press_and_release('4')
                                                    time.sleep(0.3)

                                                    pyautogui.click(x_registro, y_registro, button='left', clicks=1, interval=0.25)
                                                    time.sleep(0.3) 

                                                    keyboard.press_and_release('space')
                                                    keyboard.press_and_release('space')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('tab')
                                                    time.sleep(0.3)

                                                    keyboard.press_and_release('n')
                                                    time.sleep(0.3)

                                                    pyautogui.click(x_registado_sim, y_registado_sim, button='left', clicks=1, interval=0.25)
                                                    time.sleep(0.3) 

                                                    keyboard.press_and_release('space')
                                                    keyboard.press_and_release('space')
                                                    time.sleep(0.3)

                                                    pyautogui.click(x_anterior, y_anterior, button='left', clicks=3, interval=0.25)
                                                    time.sleep(0.1)
                                        

                                                # print("Retornando ao código principal.")
                                                break
                                        

                                    if not encontrou_linha:
                                        contador_introduzido += 1
                                    # print("Nenhuma linha correspondente encontrada. Pressionando Ctrl + I.")
                                        keyboard.press_and_release('ctrl+i')
                                        time.sleep(2)
                                        keyboard.press_and_release('tab')
                                        for _ in range(3):
                                            keyboard.press_and_release('i')
                                            time.sleep(1)
                                        for _ in range(3):    
                                            keyboard.press_and_release('tab')
                                        keyboard.write('rede perda')
                                        
                                        pyautogui.click(x_seguinte, y_seguinte, button='left', clicks=2, interval=0.25)
                                        time.sleep(1)
                                        
                                        for _ in range(4):    
                                            keyboard.press_and_release('space')
                                        time.sleep(1)
                                        
                                        pyautogui.click(x_processamento,y_processamento, button='left', clicks=2, interval=0.25)
                                        time.sleep(1)

                                        pyautogui.click(x_processamento_seguinte, y_processamento_seguinte, button='left', clicks=2, interval=0.25)
                                        time.sleep(4)

                                        pyautogui.click(x_imprimir, y_imprimir, button='left', clicks=1, interval=0.25)
                                        time.sleep(2)

                                        pyautogui.click( x_processamento_sim, y_processamento_sim, button='left', clicks=1, interval=0.25)
                                        time.sleep(4)

                                        pyautogui.click(x_cancelar, y_cancelar, button='left', clicks=1, interval=0.25)
                                        time.sleep(2)

                                        pyautogui.click(x_registado, y_registado, button='left', clicks=2, interval=0.25)
                                        time.sleep(1)

                                        pyautogui.click(x_registado_seguinte, y_registado_seguinte, button='left', clicks=2, interval=0.25)
                                        time.sleep(3)

                                        keyboard.press_and_release('space')

                                        pyautogui.click(x_registado_cmp1, y_registado_cmp1 , button='left', clicks=1, interval=0.25)
                                        time.sleep(5)

                                        keyboard.press_and_release('r')
                                        time.sleep(0.5)

                                        keyboard.press_and_release('tab')
                                        time.sleep(0.5)

                                        keyboard.press_and_release('v')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('tab')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('2')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('tab')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('2')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('shift+tab')
                                        time.sleep(2)

                                        screenshot = pyautogui.screenshot(region=area2)
                                        texto_extraido = pytesseract.image_to_string(screenshot)
                                        
                                        texto_extraido = texto_extraido.strip()  # Remove espaços no início e no fim
                                        texto_extraido = texto_extraido.replace("\n", "")  # Remove quebras de linha

                                        # Mostrar o texto extraído para depuração
                                        print("Texto extraído:\n", texto_extraido)

                                        # Comparar o texto
                                        if texto_extraido == "22000 V":
                                            keyboard.press_and_release('tab')
                                            time.sleep(0.3)
                                            keyboard.press_and_release('2')
                                            time.sleep(0.3)
                                            
                                        else:
                                            keyboard.press_and_release('tab')
                                            time.sleep(0.3)
                                            
                                        keyboard.press_and_release('tab')
                                        time.sleep(0.3)
                                        keyboard.press_and_release('s')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('tab')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('s')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('tab')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('r')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('tab')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('2')
                                        keyboard.press_and_release('0')
                                        keyboard.press_and_release('0')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('tab')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('4')
                                        keyboard.press_and_release('3')
                                        time.sleep(0.3)
                    
                                        keyboard.press_and_release('tab')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('4')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('tab')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('2')
                                        keyboard.press_and_release('0')
                                        keyboard.press_and_release('0')
                                        time.sleep(1)
                                        
                                        pyautogui.click(x_registado_cmp_potencia, y_registado_cmp1_potencia, button='left', clicks=1, interval=0.25)
                                        time.sleep(0.3)

                                        keyboard.press_and_release('2')
                                        keyboard.press_and_release('.')
                                        keyboard.press_and_release('2')
                                        time.sleep(0.3)
                                        
                                        pyautogui.click(x_propriedade, y_propriedade, button='left', clicks=2, interval=0.25)
                                        time.sleep(0.3)
                                        
                                        keyboard.press_and_release('backspace')
                                        time.sleep(0.3) 

                                        keyboard.press_and_release('1')
                                        keyboard.press_and_release('7')
                                        keyboard.press_and_release('2')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('tab')
                                        time.sleep(0.3)

                                        pyautogui.click(x_gis_x, y_gis_x, button='left', clicks=2, interval=0.25)
                                        time.sleep(0.3)
                                        
                                        keyboard.press_and_release('backspace')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('2')
                                        time.sleep(0.3)

                                        pyautogui.click(x_gis_y, y_gis_y, button='left', clicks=2, interval=0.25)
                                        time.sleep(0.3)
                                        
                                        keyboard.press_and_release('backspace')
                                        time.sleep(0.3)
                                        
                                        keyboard.press_and_release('1')
                                        time.sleep(0.3)

                                        
                                        pyautogui.click(x_estado_instalacao, y_estado_instalacao, button='left', clicks=1, interval=0.25)
                                        time.sleep(0.3) 

                                        keyboard.press_and_release('s')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('tab')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('s')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('tab')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('s')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('tab')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('s')
                                        time.sleep(0.3)
                                    
                                        pyautogui.click(x_num_luz, y_num_luz , button='left', clicks=1, interval=0.25)
                                        time.sleep(0.3) 

                                        keyboard.press_and_release('8')
                                        time.sleep(1)

                                        pyautogui.click(x_quartos, y_quartos, button='left', clicks=1, interval=0.25)
                                        time.sleep(0.3) 

                                        keyboard.press_and_release('4')
                                        time.sleep(0.3)
                                        

                                        pyautogui.click(x_registro, y_registro, button='left', clicks=1, interval=0.25)
                                        time.sleep(0.3) 

                                        keyboard.press_and_release('space')
                                        keyboard.press_and_release('space')
                                        keyboard.press_and_release('space')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('tab')
                                        time.sleep(0.3)

                                        keyboard.press_and_release('n')
                                        time.sleep(0.3)
                                        

                                        pyautogui.click( x_registado_sim, y_registado_sim, button='left', clicks=1, interval=0.25)
                                        time.sleep(0.3) 

                                        keyboard.press_and_release('space')
                                        keyboard.press_and_release('space')
                                        keyboard.press_and_release('space')
                                        time.sleep(0.3)
                                        
                                        pyautogui.click(x_anterior, y_anterior, button='left', clicks=3, interval=0.25)
                                        time.sleep(0.1)
                                        
                                        continue
            except KeyboardInterrupt:
                print("\nPrograma encerrado pelo utilizador.") 
            # Hora de término
            hora_fim = time.time() + 2 * 3600
            tempo_execucao = (hora_fim - hora_inicio)/60
            numeros_errados_totais = contador_introduzido - numeros_errados
            self.add_log(f"\n\nProcesso concluído!", "Automação")
            self.add_log(f"Hora de início: {time.strftime('%H:%M:%S', time.gmtime(hora_inicio))}", "Automação")
            self.add_log(f"Hora de término: {time.strftime('%H:%M:%S', time.gmtime(hora_fim))}", "Automação")
            self.add_log(f"Tempo total de execução: {tempo_execucao:.2f} min", "Automação")
            self.add_log(f"Número de dados para processamento: {num_dados}", "Automação")
            self.add_log(f"Total de dados processados: {total_processados}", "Automação")
            self.add_log(f"Total de inspecções introduzidas: {contador_introduzido}", "Automação")
            self.add_log(f"Total de inspecções colocadas em processamento: {contador_registados}", "Automação")
            self.add_log(f"Total de inspecções resolvidas: {contador_em_processamento}", "Automação")
            self.add_log(f"Numeros errados: {numeros},", "Automação")
            self.add_log(f"Total de numeros errados: {numeros_errados}", "Automação")
        except Exception as e:
            error_msg = f"Erro na automação: {str(e)}"
            self.add_log(error_msg, "Automação")
            messagebox.showerror("Erro", error_msg)
#Função para editar coordenadas manualmente
    def editar_coordenadas(self):
        if not coordenadas:
            messagebox.showwarning("Aviso", "Nenhuma coordenada disponível para edição.")
            self.add_log("Tentativa de edição sem coordenadas disponíveis", "Edição")
            return
        
        self.add_log("Abrindo editor de coordenadas", "Edição")
        nova_janela = tk.Toplevel(self.janela)
        nova_janela.title("Editar Coordenadas")
        
        # Label e Combobox para selecionar a chave
        tk.Label(nova_janela, text="Selecione a chave:").grid(row=0, column=0, padx=10, pady=10)
        chave_combobox = ttk.Combobox(nova_janela, values=list(coordenadas.keys()))
        chave_combobox.grid(row=0, column=1, padx=10, pady=10)
        chave_combobox.current(0)
        
        # Labels e campos de entrada para X e Y
        tk.Label(nova_janela, text="X:").grid(row=1, column=0, padx=10, pady=10)
        x_entry = tk.Entry(nova_janela)
        x_entry.grid(row=1, column=1, padx=10, pady=10)
        
        tk.Label(nova_janela, text="Y:").grid(row=2, column=0, padx=10, pady=10)
        y_entry = tk.Entry(nova_janela)
        y_entry.grid(row=2, column=1, padx=10, pady=10)
        
        # Função para carregar os valores atuais da chave selecionada
        def carregar_valores(event=None):
            chave = chave_combobox.get()
            if chave in coordenadas:
                x, y = coordenadas[chave]
                x_entry.delete(0, tk.END)
                x_entry.insert(0, str(x))
                y_entry.delete(0, tk.END)
                y_entry.insert(0, str(y))
        
        chave_combobox.bind("<<ComboboxSelected>>", carregar_valores)
        carregar_valores()
        
        # Função para adicionar uma nova chave
        def adicionar_chave():
            nova_chave = chave_combobox.get()
            if nova_chave and nova_chave not in coordenadas:
                coordenadas[nova_chave] = (0, 0)
                chave_combobox["values"] = list(coordenadas.keys())
                chave_combobox.set(nova_chave)
                carregar_valores()
                messagebox.showinfo("Sucesso", f"Chave '{nova_chave}' adicionada.")
                self.add_log(f"Chave '{nova_chave}' adicionada", "Edição")
            else:
                messagebox.showwarning("Aviso", "A chave já existe ou é inválida.")
        
        # Função para excluir a chave selecionada
        def excluir_chave():
            chave = chave_combobox.get()
            if chave in coordenadas:
                resposta = messagebox.askyesno("Confirmar", f"Tem certeza que deseja excluir a chave '{chave}'?")
                if resposta:
                    del coordenadas[chave]
                    chave_combobox["values"] = list(coordenadas.keys())
                    if coordenadas:
                        chave_combobox.current(0)
                        carregar_valores()
                    else:
                        chave_combobox.set("")
                        x_entry.delete(0, tk.END)
                        y_entry.delete(0, tk.END)
                    messagebox.showinfo("Sucesso", f"Chave '{chave}' excluída.")
                    self.add_log(f"Chave '{chave}' excluída", "Edição")
        
        # Função para salvar a edição
        def salvar_edicao():
            chave = chave_combobox.get()
            try:
                x = int(x_entry.get())
                y = int(y_entry.get())
                coordenadas[chave] = (x, y)
                messagebox.showinfo("Sucesso", f"Coordenada {chave} atualizada para ({x}, {y}).")
                self.salvar_coordenadas()
                nova_janela.destroy()
                self.add_log(f"Coordenada {chave} atualizada para ({x}, {y})", "Edição")
            except ValueError:
                messagebox.showerror("Erro", "Valores de X e Y devem ser números inteiros.")
        
        # Botões para adicionar, excluir e salvar
        tk.Button(nova_janela, text="Adicionar Chave", command=adicionar_chave).grid(row=3, column=0, pady=10)
        tk.Button(nova_janela, text="Excluir Chave", command=excluir_chave).grid(row=3, column=1, pady=10)
        tk.Button(nova_janela, text="Salvar", command=salvar_edicao).grid(row=4, column=0, columnspan=2, pady=10)
#função para carregar o arquivo
    def selecionar_arquivo_txt(self):
            caminho = filedialog.askopenfilename(title="Selecionar arquivo", filetypes=[("Arquivos de texto", "*.txt")])
            if caminho:
                self.dados_txt = self.carregar_dados_txt(caminho)
                self.label_arquivo_txt.config(text=f"Arquivo carregado: {os.path.basename(caminho)}", fg="green")
#função para lidar com os dados do arquivo
    def carregar_dados_txt(self, caminho):
        global dados
        dados = []
        try:
            with open(caminho, 'r', encoding='utf-8') as arquivo:
                linhas = [linha.strip() for linha in arquivo if linha.strip()]
                for i in range(0, len(linhas), 2):
                    nome = linhas[i]
                    if i+1 < len(linhas) and linhas[i+1].startswith("Contador:"):
                        contador = linhas[i+1].replace("Contador:", "").strip()
                        dados.append(contador)
            self.add_log(f"{len(dados)} registros carregados do arquivo", "Sucesso")
            self.add_log(dados)
            return dados
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar dados: {e}")
            self.add_log(f"Erro ao carregar dados: {e}", "Erro")
            return []
# Iniciar processo
if __name__ == "__main__":
    app = Aplicacao()
    app.janela.mainloop()
