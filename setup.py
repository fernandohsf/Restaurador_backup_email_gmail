from cx_Freeze import setup, Executable
import os

caminhoExe = os.path.dirname(os.path.realpath(__file__))

build_exe_options = {
    "packages": ["os", "sys", "time", "fitz", "shutil", "pdfkit", "mailbox", "email", "icalendar", "tkinter", "threading", "bs4", "PIL", "win32com", "zipfile"],
    "include_files":[f"{caminhoExe}\\imagens", f"{caminhoExe}\\wkhtmltopdf"],
    "excludes": [],
}

setup(
    name="Recuperador de e-mail para pdf",
    version="1.0",
    description="Extrai os e-mail de um arquivo de backup .mbox e converte em pdf",
    options={"build_exe": build_exe_options},
    executables=[Executable("tela_principal.pyw", base="Win32GUI", target_name="Restaurador de backup de e-mail")],
)