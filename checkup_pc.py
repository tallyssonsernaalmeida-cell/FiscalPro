import psutil
import platform
import os
from datetime import datetime

def bytes_to_gb(bytes):
    return round(bytes / (1024**3), 2)

print("=== CHECKUP DO PC ===")
print(f"Sistema Operacional: {platform.system()} {platform.release()}")
print(f"Nome do computador: {platform.node()}")
print(f"Processador: {platform.processor()}")
print(f"Arquitetura: {platform.architecture()[0]}")
print(f"Hora: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
print("")

# CPU
print("=== CPU ===")
print(f"Núcleos físicos: {psutil.cpu_count(logical=False)}")
print(f"Núcleos lógicos: {psutil.cpu_count(logical=True)}")
print(f"Uso atual: {psutil.cpu_percent(interval=1)}%")
print("")

# Memória RAM
mem = psutil.virtual_memory()
print("=== MEMÓRIA RAM ===")
print(f"Total: {bytes_to_gb(mem.total)} GB")
print(f"Disponível: {bytes_to_gb(mem.available)} GB")
print(f"Uso: {mem.percent}%")
print("")

# Disco
print("=== DISCO ===")
for part in psutil.disk_partitions():
    usage = psutil.disk_usage(part.mountpoint)
    print(f"{part.device} ({part.mountpoint}) - Total: {bytes_to_gb(usage.total)} GB, Uso: {usage.percent}%")

print("")

# Rede
net = psutil.net_io_counters()
print("=== REDE ===")
print(f"Bytes enviados: {bytes_to_gb(net.bytes_sent)} GB")
print(f"Bytes recebidos: {bytes_to_gb(net.bytes_recv)} GB")