# gunicorn.conf.py

# Socket
bind = "0.0.0.0:5231"

# Workers
workers = 4  # Ajuste conforme os Cores da sua VPS
worker_class = "gthread"
threads = 2

# Logging
accesslog = "-"  # Envia logs de acesso para o stdout
errorlog = "-"   # Envia logs de erro para o stdout
loglevel = "info"

# Timeout
timeout = 120  # Aumentado para 120s devido ao processamento do Gemini
keepalive = 5
