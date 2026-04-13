import sys
import subprocess

print("Установка зависимостей...")
result = subprocess.run([sys.executable, "install.py"])

if result.returncode != 0:
    print("Ошибка при установке зависимостей. Запуск отменён.")
    sys.exit(1)

print("Запуск приложения...")
subprocess.run([sys.executable, "main.py"])
