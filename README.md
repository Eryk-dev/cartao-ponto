# 🐳 API de Cartão-Ponto (Flask + Docker)

## 🎯 Visão Geral
API em Flask para processar cartões-ponto, otimizada para deploy em VPS com Docker, Gunicorn e Nginx.

## 🚀 Deploy Rápido com Docker

### 1. Pré-requisitos
- Docker e Docker Compose instalados no seu VPS.

### 2. Configure a Chave da API
- Renomeie `.env.example` para `.env`:
  ```bash
  mv .env.example .env
  ```
- Edite o arquivo `.env` e insira sua chave da API do Google Gemini:
  ```
  GEMINI_API_KEY=SUA_CHAVE_API_AQUI
  ```

### 3. Suba os Contêineres
- Execute o Docker Compose em modo detached (-d):
  ```bash
  docker-compose up --build -d
  ```

### 4. Teste a API
- A API estará disponível na porta 80 do seu VPS.
  ```bash
  # Verificar saúde
  curl http://SEU_IP_DO_VPS/health

  # Processar um mês (com as imagens de teste)
  curl -X POST -F "file1=@quinzena1.jpg" -F "file2=@quinzena2.jpg" http://SEU_IP_DO_VPS/process-month
  ```

## ⚙️ Gerenciando os Contêineres
- **Ver Logs:**
  ```bash
  docker-compose logs -f api  # Logs da API
  docker-compose logs -f nginx # Logs do Nginx
  ```
- **Parar Contêineres:**
  ```bash
  docker-compose down
  ```

## 📋 Endpoint Principal
| Método | Endpoint | Descrição |
|--------|----------|-----------|
| POST | `/process-month` | Processa duas imagens de quinzena e retorna o Excel em base64. |

## 📊 Exemplo de Resposta
```json
{
  "success": true,
  "message": "Mês completo processado com sucesso! ✅",
  "file_base64": "UEsDBBQAAAAIA...",
  "time_entries": {
    "16": { "times": ["09:04", "13:00", "14:00", "18:02"], "painted": false },
    "17": { "times": ["08:56", "13:00", "14:00", "18:01"], "painted": true }
  },
  "summary": {
    "output_file": "Registro_Ponto_20250730_145940.xlsx",
    ...
  }
}
```

## 🔧 Estrutura do Projeto
```
cartao-ponto-flask/
├── app.py                    # API Flask
├── requirements.txt          # Dependências Python
├── Dockerfile                # Build da imagem da API
├── docker-compose.yml        # Orquestração (API + Nginx)
├── gunicorn.conf.py          # Configurações do Gunicorn
├── nginx.conf                # Configurações do Nginx
├── .env.example              # Exemplo de variáveis de ambiente
├── .dockerignore             # Arquivos ignorados pelo Docker
├── Registro Ponto.XLSX       # Template Excel
├── quinzena1.jpg             # Imagem de teste
├── quinzena2.jpg             # Imagem de teste
└── README.md                 # Esta documentação
```
