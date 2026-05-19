import pymssql

print("="*50)
print("TESTE DE CONEXAO SQL SERVER")
print("="*50)

# Configuracoes
server = "DESKTOP-74TBS68"
database = "IAF_BD"
username = "ControlDesk"
password = "Control@2026"

print(f"\nTentando conectar com:")
print(f"  Servidor: {server}")
print(f"  Banco: {database}")
print(f"  Usuario: {username}")
print(f"  Senha: {'*' * len(password)}")

# Tentativa 1: Sem porta especifica
try:
    print("\n[1] Tentando conexao padrao...")
    conn = pymssql.connect(
        server=server,
        database=database,
        user=username,
        password=password
    )
    print("SUCESSO! Conexao estabelecida!")
    conn.close()
except Exception as e:
    print(f"FALHA: {e}")

# Tentativa 2: Com porta 1433
try:
    print("\n[2] Tentando com porta 1433...")
    conn = pymssql.connect(
        server=f"{server}:1433",
        database=database,
        user=username,
        password=password
    )
    print("SUCESSO! Conexao estabelecida!")
    conn.close()
except Exception as e:
    print(f"FALHA: {e}")

# Tentativa 3: Autenticacao Windows
try:
    print("\n[3] Tentando com autenticacao Windows...")
    conn = pymssql.connect(
        server=server,
        database=database,
        trusted=True
    )
    print("SUCESSO! Conexao estabelecida com Windows Auth!")
    conn.close()
except Exception as e:
    print(f"FALHA: {e}")

print("\n" + "="*50)
print("TESTE CONCLUIDO")
print("="*50)