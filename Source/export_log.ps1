# Configurar o caminho de saída para o log
$logPath = "caminho\para\seu\arquivo"

# Especificar o nome do log e o filtro de eventos
$logName = "Microsoft-Windows-PrintService/Operational"
$eventID = 307  # Exemplo: Evento com ID 1000


# Filtrar eventos específicos
$events = Get-WinEvent -LogName $logName -FilterXPath "*[System[(EventID=$eventID)]]"

# Selecionar propriedades
$eventProperties = $events | Select-Object -Property TimeCreated, id, Message

# Exportar os eventos para um arquivo de texto
$eventProperties | Export-Csv -Path $logPath -NoTypeInformation -Encoding utf8


# Opcional: Adicionar data e hora ao log
Add-Content -Path $logPath -Value ("`nLog exportado em: " + (Get-Date))