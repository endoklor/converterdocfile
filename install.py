import sys
import subprocess
import importlib.util

PACKAGES = {
    "customtkinter": "customtkinter",
    "docx2pdf": "docx2pdf",
}

def is_installed(import_name):
    return importlib.util.find_spec(import_name) is not None

def install(package_name):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package_name])

for import_name, package_name in PACKAGES.items():
    if is_installed(import_name):
        print(f"[OK] {package_name} уже установлен")
    else:
        print(f"[...] Устанавливается {package_name}...")
        try:
            install(package_name)
            print(f"[OK] {package_name} успешно установлен")
        except subprocess.CalledProcessError:
            print(f"[ERR] Не удалось установить {package_name}")
