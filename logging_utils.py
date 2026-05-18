"""Logger compartilhado pelas automações do SIFAMA."""

import os
from datetime import datetime


def _ts() -> str:
    """Só devolve o horário no formato usado nos logs."""
    return datetime.now().strftime("%H:%M:%S")


class Logger:
    """Centraliza os logs da automação na tela e, se quiser, em arquivo."""

    def __init__(self, log_callback=None):
        self.logs = []
        self.log_callback = log_callback
        self.log_file = None

    def set_log_file(self, path):
        """Liga a gravação em arquivo e cria a pasta se ela ainda não existir."""
        self.log_file = path
        if path:
            dir_log = os.path.dirname(path)
            if dir_log and not os.path.isdir(dir_log):
                try:
                    os.makedirs(dir_log, exist_ok=True)
                except Exception:
                    self.log_file = None

    def log(self, mensagem, tipo="INFO"):
        """Registra uma mensagem no mesmo formato usado no restante da automação."""
        log_entry = f"[{_ts()}] [{tipo}] {mensagem}"
        self.logs.append(log_entry)

        if self.log_callback:
            self.log_callback(log_entry, tipo)

        if self.log_file:
            try:
                with open(self.log_file, "a", encoding="utf-8") as arquivo_log:
                    arquivo_log.write(log_entry + "\n")
            except Exception:
                pass

        print(log_entry)

    def get_logs(self):
        """Devolve o histórico inteiro em texto corrido."""
        return "\n".join(self.logs)
