"""Ponto de entrada da automação Certronic → Matriz (IN HAUS)."""

import shutil
import sys
from pathlib import Path

from src.gui import App


def _dir_base() -> Path:
    """Pasta onde fica o config.yaml editável.

    - Empacotado (PyInstaller --onefile): a pasta do próprio .exe. Assim as
      configurações que a operação salva na aba Configurações PERSISTEM entre
      execuções (o _MEIPASS é temporário e some a cada run).
    - Rodando como script (desenvolvimento): a pasta do projeto.
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent


if __name__ == "__main__":
    base = _dir_base()
    config = Path(sys.argv[1]) if len(sys.argv) > 1 else base / "config.yaml"

    # Primeira execução do .exe: se ainda não existe um config ao lado do
    # executável, cria a partir do modelo embutido no pacote (preserva o
    # de-para de documentos e as regras). A partir daí, a aba Configurações
    # edita este arquivo, que persiste.
    if not config.exists():
        modelo = Path(getattr(sys, "_MEIPASS", base)) / "config.yaml"
        if modelo.exists() and modelo.resolve() != config.resolve():
            shutil.copy(modelo, config)

    App(config).executar()
