import win32com.client

# Inicializa Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Acessa a Caixa de Saída (Outbox)
outbox = outlook.GetDefaultFolder(4)  # 4 = olFolderOutbox

# Conta quantos e-mails estão lá
itens = outbox.Items
print(f"📬 E-mails encontrados na Caixa de Saída: {len(itens)}")

# Exclui todos
for item in list(itens):
    try:
        assunto = item.Subject
        item.Delete()
        print(f"🗑️ Email removido: {assunto}")
    except Exception as e:
        print(f"⚠️ Erro ao excluir item: {e}")

print("✅ Limpeza da caixa de saída concluída.")
