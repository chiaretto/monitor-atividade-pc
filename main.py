"""
Monitor de Atividade no Windows
--------------------------------
Open Source - Código simplificado e comentado para evoluções futuras.

Funcionalidades:
- Registra períodos de atividade/ociosidade do usuário.
- Registra programas e documentos em foco.
- Exibe uma linha do tempo diária em uma interface Tkinter.
- Minimiza para a system tray.

Evoluções futuras possíveis:
- Suporte multiusuário.
- Exportação de relatórios em CSV/Excel.
- Dashboard web em vez de Tkinter.
- Suporte multiplataforma (Linux/Mac).
"""

import ctypes
import sqlite3
import threading
import time
from datetime import datetime, timedelta
import tkinter as tk
from tkcalendar import DateEntry
import pystray
from PIL import Image, ImageDraw
import win32gui
import win32process
import psutil

# ===================== Configurações =====================
LIMITE_OCIOSO = 300  # Segundos de inatividade até considerar "Ocioso"

# ===================== Banco de Dados =====================
class BancoUtil:
    """
    Classe utilitária para abstrair o SQLite.
    Usa lock para evitar problemas em múltiplas threads.
    """
    _lock = threading.Lock()

    def __init__(self, db_path="atividade.db"):
        try:
            self.conn = sqlite3.connect(db_path, check_same_thread=False, timeout=10)
            self.cursor = self.conn.cursor()
            self.criar_tabelas()
        except sqlite3.Error as e:
            print(f"[ERRO] Falha ao inicializar banco de dados: {e}")
            raise

    def criar_tabelas(self):
        """Cria tabelas necessárias (se não existirem)."""
        with BancoUtil._lock:
            self.cursor.execute("""
                CREATE TABLE IF NOT EXISTS log_atividade (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    status TEXT,
                    data_hora TEXT,
                    data_hora_final TEXT
                )
            """)
            self.cursor.execute("""
                CREATE TABLE IF NOT EXISTS programas_documentos (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    nome_programa TEXT,
                    nome_documento TEXT,
                    data_hora_inicio TEXT,
                    data_hora_final TEXT
                )
            """)
            self.conn.commit()

    def executar(self, query, params=(), fetch=False):
        """Executa query no banco, com tratamento de erro."""
        try:
            with BancoUtil._lock:
                self.cursor.execute(query, params)
                self.conn.commit()
                if fetch:
                    return self.cursor.fetchall()
                return self.cursor.lastrowid
        except sqlite3.Error as e:
            print(f"[ERRO] Query falhou: {e} | SQL: {query} | Params: {params}")
            return None

# ===================== Captura Idle Windows =====================
class LASTINPUTINFO(ctypes.Structure):
    _fields_ = [("cbSize", ctypes.c_uint), ("dwTime", ctypes.c_uint)]

def get_idle_time():
    """Retorna tempo (segundos) desde a última interação do usuário."""
    try:
        lii = LASTINPUTINFO()
        lii.cbSize = ctypes.sizeof(LASTINPUTINFO)
        ctypes.windll.user32.GetLastInputInfo(ctypes.byref(lii))
        millis = ctypes.windll.kernel32.GetTickCount() - lii.dwTime
        return millis / 1000
    except Exception as e:
        print(f"[ERRO] Falha ao capturar idle time: {e}")
        return 0

# ===================== Tooltip =====================
class ToolTip:
    """Exibe tooltip ao passar o mouse sobre o canvas."""
    def __init__(self, widget):
        self.widget = widget
        self.tipwindow = None

    def showtip(self, text, x, y):
        self.hidetip()
        try:
            self.tipwindow = tw = tk.Toplevel(self.widget)
            tw.wm_overrideredirect(True)
            tw.wm_geometry(f"+{x}+{y}")
            label = tk.Label(
                tw, text=text,
                background="#ffffe0", relief="solid",
                borderwidth=1, font=("Arial", 8)
            )
            label.pack()
        except Exception as e:
            print(f"[ERRO] Falha ao exibir tooltip: {e}")

    def hidetip(self):
        if self.tipwindow:
            self.tipwindow.destroy()
            self.tipwindow = None

# ===================== Log de Atividade =====================
class LogAtividade:
    """Classe principal que integra captura, interface e persistência."""
    def __init__(self, limite_ocioso=LIMITE_OCIOSO, intervalo_cheque=5):
        self.db = BancoUtil()
        self.status_atual = None
        self.limite_ocioso = limite_ocioso
        self.intervalo_cheque = intervalo_cheque
        self.running = True

        # Controle de aplicativos/documentos
        self.ultimo_app = None
        self.ultimo_titulo = None
        self.inicio_app = None

        # Data padrão selecionada
        self.data_selecionada = datetime.now().date()

        # Cria interface
        self.desenhar_interface()

        # Configura tray
        self.root.protocol("WM_DELETE_WINDOW", self.minimizar_tray)
        self.tray_icon = pystray.Icon(
            "atividade_usuario",
            self.criar_icone_tray(),
            "Atividade Usuário",
            menu=pystray.Menu(
                pystray.MenuItem("Abrir", self.restaurar_janela),
                pystray.MenuItem("Sair", self.sair)
            )
        )
        threading.Thread(target=self.tray_icon.run, daemon=True).start()

        # Inicia monitoramento em thread separada
        self.monitor_thread = threading.Thread(target=self.monitorar, daemon=True)
        self.monitor_thread.start()

        # Atualização periódica da timeline
        self.root.after(1000, self.atualizar_canvas)
        self.root.mainloop()

    # ===================== Interface =====================
    def desenhar_interface(self):
        """Cria a interface Tkinter."""
        self.canvas_height = 100
        self.root = tk.Tk()
        self.root.title("Linha do Tempo - Atividade")

        # Navegação de datas
        frame_datas = tk.Frame(self.root)
        frame_datas.pack(pady=5)

        tk.Button(frame_datas, text="◀ Anterior", command=self.dia_anterior).pack(side=tk.LEFT, padx=5)
        self.cal = DateEntry(frame_datas, width=12,
                             background='darkblue', foreground='white', borderwidth=2,
                             year=self.data_selecionada.year,
                             month=self.data_selecionada.month,
                             day=self.data_selecionada.day)
        self.cal.pack(side=tk.LEFT, padx=5)
        self.cal.bind("<<DateEntrySelected>>", self.on_data_selecionada)
        tk.Button(frame_datas, text="Próximo ▶", command=self.dia_proximo).pack(side=tk.LEFT, padx=5)
        tk.Button(frame_datas, text="Hoje", command=self.dia_hoje).pack(side=tk.LEFT, padx=5)

        # Label de status
        self.label_status = tk.Label(self.root, text="", font=("Arial", 10))
        self.label_status.pack()

        # Canvas timeline
        self.canvas = tk.Canvas(self.root, width=1440, height=self.canvas_height, bg="white")
        self.canvas.pack()
        self.rects = [self.canvas.create_rectangle(x, 0, x+1, self.canvas_height, fill="gray", outline="") for x in range(1440)]

        # Tooltip
        self.tooltip = ToolTip(self.canvas)
        self.canvas.bind("<Motion>", self.on_mouse_move)
        self.canvas.bind("<Leave>", lambda e: self.tooltip.hidetip())

    def criar_icone_tray(self):
        """Cria ícone verde simples para a tray."""
        img = Image.new("RGB", (64, 64), color="white")
        d = ImageDraw.Draw(img)
        d.rectangle([16, 16, 48, 48], fill="green")
        return img

    # ===================== Navegação de dias =====================
    def dia_hoje(self): self.set_data(datetime.now().date())
    def dia_anterior(self): self.set_data(self.data_selecionada - timedelta(days=1))
    def dia_proximo(self): self.set_data(self.data_selecionada + timedelta(days=1))
    def set_data(self, nova_data):
        self.data_selecionada = nova_data
        self.cal.set_date(nova_data)
        self.atualizar_canvas()

    # ===================== Tooltip Canvas =====================
    def on_mouse_move(self, event):
        x = event.x
        if 0 <= x < 1440:
            hora, minuto = divmod(x, 60)
            self.tooltip.showtip(f"{hora:02d}:{minuto:02d}", event.x_root + 10, event.y_root + 10)
        else:
            self.tooltip.hidetip()

    def on_data_selecionada(self, event):
        self.data_selecionada = self.cal.get_date()
        self.atualizar_canvas()

    # ===================== Registro de Status =====================
    def registrar_status(self, novo_status):
        agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        if self.status_atual != novo_status:
            if self.status_atual:
                self.db.executar("""
                    UPDATE log_atividade
                    SET data_hora_final = ?
                    WHERE status = ? AND data_hora_final IS NULL
                """, (agora, self.status_atual))
            self.db.executar("""
                INSERT INTO log_atividade (status, data_hora)
                VALUES (?, ?)
            """, (novo_status, agora))
            self.status_atual = novo_status

    # ===================== Timeline =====================
    def atualizar_canvas(self):
        """
        Atualiza a linha do tempo e os totais de horas.
        - Calcula horas ativas, ociosas e total do dia selecionado.
        - Atualiza as cores do canvas.
        - Atualiza label de status.
        """
        hoje = self.data_selecionada

        # ===================== Consulta de horas totais =====================
        try:
            row = self.db.executar("""
                SELECT
                    printf('%02d:%02d',
                        CAST(SUM((julianday(COALESCE(data_hora_final, DATETIME('now', 'localtime'))) - julianday(data_hora)) * 24) AS INTEGER),
                        CAST((SUM((julianday(COALESCE(data_hora_final, DATETIME('now', 'localtime'))) - julianday(data_hora)) * 24) * 60) % 60 AS INTEGER)
                    ) AS horas_totais,

                    printf('%02d:%02d',
                        CAST(SUM(CASE WHEN status = 'Ativo' THEN (julianday(COALESCE(data_hora_final, DATETIME('now', 'localtime'))) - julianday(data_hora)) * 24 ELSE 0 END) AS INTEGER),
                        CAST((SUM(CASE WHEN status = 'Ativo' THEN (julianday(COALESCE(data_hora_final, DATETIME('now', 'localtime'))) - julianday(data_hora)) * 24 ELSE 0 END) * 60) % 60 AS INTEGER)
                    ) AS horas_ativas,

                    printf('%02d:%02d',
                        CAST(SUM(CASE WHEN status = 'Ocioso' THEN (julianday(COALESCE(data_hora_final, DATETIME('now', 'localtime'))) - julianday(data_hora)) * 24 ELSE 0 END) AS INTEGER),
                        CAST((SUM(CASE WHEN status = 'Ocioso' THEN (julianday(COALESCE(data_hora_final, DATETIME('now', 'localtime'))) - julianday(data_hora)) * 24 ELSE 0 END) * 60) % 60 AS INTEGER)
                    ) AS horas_ociosas,

                    ROUND(
                        100.0 * SUM(CASE WHEN status = 'Ocioso' THEN (julianday(COALESCE(data_hora_final, DATETIME('now', 'localtime'))) - julianday(data_hora)) * 24 ELSE 0 END)
                        / SUM((julianday(COALESCE(data_hora_final, DATETIME('now', 'localtime'))) - julianday(data_hora)) * 24), 2
                    ) AS percentual_horas_ociosas
                FROM log_atividade
                WHERE date(data_hora) = ?
            """, (hoje.strftime("%Y-%m-%d"), ), fetch=True)

            if row and len(row) > 0:
                horas_totais, horas_ativas, horas_ociosas, pct_ociosas = row[0]
            else:
                horas_totais = horas_ativas = horas_ociosas = "00:00"
                pct_ociosas = 0

        except Exception as e:
            # Evita crash se a query falhar
            print(f"[ERRO] Falha ao calcular horas: {e}")
            horas_totais = horas_ativas = horas_ociosas = "00:00"
            pct_ociosas = 0

        # ===================== Atualiza timeline visual =====================
        try:
            rows = self.db.executar("""
                SELECT status, data_hora, COALESCE(data_hora_final, DATETIME('now', 'localtime')) AS data_hora_final
                FROM log_atividade
            """, fetch=True) or []

            # Inicializa cores (cinza = sem atividade)
            colors = ["gray"] * 1440

            for status, dh_inicio, dh_fim in rows:
                inicio = datetime.strptime(dh_inicio, "%Y-%m-%d %H:%M:%S")
                fim = datetime.strptime(dh_fim, "%Y-%m-%d %H:%M:%S")
                if inicio.date() != hoje:
                    continue

                min_inicio = inicio.hour * 60 + inicio.minute
                min_fim = fim.hour * 60 + fim.minute
                cor = "green" if status == "Ativo" else "red"

                for i in range(min_inicio, min_fim + 1):
                    if 0 <= i < 1440:
                        colors[i] = cor

            # Azul claro 08h-18h para blocos cinza
            for minuto in range(8*60, 18*60):
                if colors[minuto] == "gray":
                    colors[minuto] = "lightblue"

            # Atualiza canvas
            for x, color in enumerate(colors):
                self.canvas.itemconfig(self.rects[x], fill=color)

        except Exception as e:
            print(f"[ERRO] Falha ao atualizar canvas: {e}")

        # ===================== Atualiza label de status =====================
        self.label_status.config(
            text=f"Total: {horas_totais} - Ativo: {horas_ativas} - Ocioso: {horas_ociosas} ({pct_ociosas}% ocioso)"
        )

        # Próxima atualização
        self.root.after(1000, self.atualizar_canvas)

    # ===================== Monitoramento =====================
    def get_focused_app(self):
        try:
            hwnd = win32gui.GetForegroundWindow()
            if hwnd == 0: return None, None
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            processo = psutil.Process(pid)
            return processo.name(), win32gui.GetWindowText(hwnd)
        except Exception:
            return None, None

    def registrar_programa(self, app, titulo):
        agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with BancoUtil._lock:
            cursor = self.db.conn.cursor()
            if self.ultimo_app and self.inicio_app:
                cursor.execute("""
                    UPDATE programas_documentos
                    SET data_hora_final = ?
                    WHERE nome_programa = ? AND nome_documento = ? AND data_hora_inicio = ? AND data_hora_final IS NULL
                """, (agora, self.ultimo_app, self.ultimo_titulo, self.inicio_app))
            cursor.execute("""
                INSERT INTO programas_documentos (nome_programa, nome_documento, data_hora_inicio)
                VALUES (?, ?, ?)
            """, (app, titulo, agora))
            self.db.conn.commit()
        self.ultimo_app, self.ultimo_titulo, self.inicio_app = app, titulo, agora
        print(f"[{datetime.now().strftime('%H:%M:%S')}] Foco mudou para: {app} - {titulo}")

    def monitorar(self):
        while self.running:
            status = "Ocioso" if get_idle_time() >= self.limite_ocioso else "Ativo"
            self.registrar_status(status)
            app, titulo = self.get_focused_app()
            if app != self.ultimo_app or titulo != self.ultimo_titulo:
                self.registrar_programa(app, titulo)
            time.sleep(self.intervalo_cheque)

    # ===================== Tray =====================
    def minimizar_tray(self): self.root.withdraw()
    def restaurar_janela(self, icon, item): self.root.deiconify()
    def sair(self, icon, item):
        self.running = False
        self.tray_icon.stop()
        self.root.destroy()

# ===================== Executa =====================
if __name__ == "__main__":
    LogAtividade()
