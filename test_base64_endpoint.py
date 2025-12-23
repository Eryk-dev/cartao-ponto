#!/usr/bin/env python3
"""
Script para testar o novo endpoint /process-month-base64
"""

import requests
import base64
import json
from datetime import datetime

def convert_image_to_base64(image_path):
    """Converte imagem para base64"""
    try:
        with open(image_path, 'rb') as image_file:
            return base64.b64encode(image_file.read()).decode('utf-8')
    except Exception as e:
        print(f"Erro ao converter {image_path}: {e}")
        return None

def test_base64_endpoint():
    """Testa o endpoint /process-month-base64"""
    
    # URL da API
    url = "http://localhost:5001/process-month-base64"
    
    print("🔄 Convertendo imagens para base64...")
    
    # Converte as imagens para base64
    image1_b64 = convert_image_to_base64('quinzena1.jpg')
    image2_b64 = convert_image_to_base64('quinzena2.jpg')
    
    if not image1_b64 or not image2_b64:
        print("❌ Erro ao converter imagens para base64")
        return
    
    print(f"✅ Imagem 1: {len(image1_b64):,} chars")
    print(f"✅ Imagem 2: {len(image2_b64):,} chars")
    
    # Payload JSON
    payload = {
        "image1_base64": image1_b64,
        "image2_base64": image2_b64
    }
    
    # Headers
    headers = {
        'Content-Type': 'application/json'
    }
    
    try:
        print("\n🚀 Enviando requisição para /process-month-base64...")
        start_time = datetime.now()
        
        response = requests.post(url, json=payload, headers=headers, timeout=180)
        
        end_time = datetime.now()
        duration = (end_time - start_time).total_seconds()
        
        print(f"⏱️  Tempo de resposta: {duration:.1f}s")
        print(f"📊 Status Code: {response.status_code}")
        
        if response.status_code == 200:
            result = response.json()
            
            print("✅ Sucesso!")
            print(f"📋 Dias processados: {result['summary']['total_days']}")
            print(f"📊 Entradas totais: {result['summary']['total_entries']}")
            print(f"📄 Excel base64: {len(result.get('file_base64', '')):.0f} chars")
            print(f"🖼️  Imagem inserida: {result['summary']['image_inserted']}")
            print(f"📝 Método: {result['summary']['input_method']}")
            
        else:
            print(f"❌ Erro {response.status_code}")
            try:
                error_data = response.json()
                print(f"Erro: {error_data.get('error', 'Erro desconhecido')}")
            except:
                print(f"Resposta: {response.text}")
                
    except requests.exceptions.Timeout:
        print("⏰ Timeout - A API demorou mais de 3 minutos para responder")
    except Exception as e:
        print(f"❌ Erro na requisição: {str(e)}")

if __name__ == "__main__":
    test_base64_endpoint()