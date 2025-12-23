# Usa uma imagem base oficial do Python
FROM python:3.9-slim

# Define o diretório de trabalho no contêiner
WORKDIR /app

# Copia os arquivos de requisitos para o diretório de trabalho
COPY requirements.txt .

# Instala as dependências
RUN pip install --no-cache-dir -r requirements.txt

# Copia o resto do código da aplicação para o diretório de trabalho
COPY . .

# Expõe a porta que o Gunicorn irá rodar
EXPOSE 5231

# Comando para rodar a aplicação usando Gunicorn
CMD ["gunicorn", "-c", "gunicorn.conf.py", "app:app"]
