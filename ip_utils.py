"""
Utilitários para obter informações de IP e navegador do usuário
"""
import streamlit as st
import requests

def get_client_ip() -> str:
    """Obtém o IP do cliente através do Streamlit"""
    try:
        # Streamlit não expõe diretamente o IP, então usamos serviços externos
        # como fallback para desenvolvimento
        try:
            # Tenta obter IP local primeiro (para desenvolvimento)
            response = requests.get('https://httpbin.org/ip', timeout=5)
            if response.status_code == 200:
                return response.json().get('origin', '127.0.0.1')
        except:
            pass
        
        # Fallback para IP local
        return '127.0.0.1'
    except Exception:
        return 'Unknown'

def get_user_agent() -> str:
    """Obtém informações do navegador do usuário"""
    try:
        # Streamlit não expõe user-agent diretamente
        # Para uma implementação mais robusta, seria necessário
        # usar headers HTTP customizados
        return st.get_option("browser.gatherUsageStats") and "Streamlit App" or "Unknown Browser"
    except Exception:
        return "Unknown Browser"

def get_client_info() -> dict:
    """Obtém informações completas do cliente"""
    return {
        'ip': get_client_ip(),
        'user_agent': get_user_agent(),
        'session_id': st.session_state.get('session_id', 'unknown')
    }
