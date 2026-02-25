"""
Configuração do Firebase para monitoramento de acessos
"""
import os
import json
from datetime import datetime, timezone, timedelta
from typing import Dict, Any, Optional

try:
    import firebase_admin
    from firebase_admin import credentials, db
    FIREBASE_AVAILABLE = True
except ImportError:
    FIREBASE_AVAILABLE = False

class FirebaseManager:
    """Gerenciador do Firebase para monitoramento de acessos"""
    
    def __init__(self):
        self.app = None
        self.initialized = False
        self.firebase_connected = False
        
    def initialize(self, firebase_config: Dict[str, Any] = None):
        """Inicializa a conexão com o Firebase"""
        if not FIREBASE_AVAILABLE:
            raise ImportError("firebase-admin não está instalado")
        
        if self.initialized:
            return
            
        try:
            # Tentar conectar ao Firebase
            if firebase_config is None:
                firebase_config = self._load_config_from_file()
            
            # Inicializa o Firebase com credenciais do Service Account
            if not firebase_admin._apps:
                cred = credentials.Certificate(firebase_config)
                firebase_admin.initialize_app(cred, {
                    'databaseURL': firebase_config['databaseURL']
                })
            
            self.app = firebase_admin.get_app()
            self.firebase_connected = True
            print("✅ Firebase conectado com sucesso!")
            
        except Exception as e:
            print(f"⚠️ Firebase não disponível: {e}")
            print("📁 Usando sistema local para melhor performance")
            self.firebase_connected = False
        
        self.initialized = True
    
    def _load_config_from_file(self) -> Dict[str, Any]:
        """Carrega configuração do Firebase de arquivo"""
        config_file = "firebase_config.json"
        
        if not os.path.exists(config_file):
            raise FileNotFoundError(f"Arquivo {config_file} não encontrado")
        
        with open(config_file, 'r', encoding='utf-8') as f:
            return json.load(f)
    
    def log_access(self, usuario: str, ip: str, user_agent: str = None) -> str:
        """Registra acesso no sistema"""
        if not self.initialized:
            raise Exception("Sistema não foi inicializado")
        
        access_data = {
            'usuario': usuario,
            'ip': ip,
            'user_agent': user_agent or 'Unknown',
            'timestamp': datetime.now(timezone(timedelta(hours=-3))).isoformat(),
            'data_hora': datetime.now(timezone(timedelta(hours=-3))).strftime('%d/%m/%Y %H:%M:%S')
        }
        
        # Sempre salvar localmente (rápido)
        self._save_local_log(access_data)
        
        # Tentar salvar no Firebase se conectado
        if self.firebase_connected:
            try:
                ref = db.reference('access_logs')
                new_ref = ref.push(access_data)
                return new_ref.key
            except Exception as e:
                print(f"⚠️ Firebase temporariamente indisponível: {e}")
                self.firebase_connected = False
        
        return f"local_{datetime.now(timezone(timedelta(hours=-3))).timestamp()}"
    
    def _save_local_log(self, access_data: Dict[str, Any]):
        """Salva log localmente"""
        try:
            log_file = "local_access_log.json"
            
            # Carregar logs existentes
            logs = []
            if os.path.exists(log_file):
                with open(log_file, 'r', encoding='utf-8') as f:
                    logs = json.load(f)
            
            # Adicionar novo log
            logs.append(access_data)
            
            # Salvar de volta
            with open(log_file, 'w', encoding='utf-8') as f:
                json.dump(logs, f, ensure_ascii=False, indent=2)
                
        except Exception as e:
            print(f"Erro ao salvar log local: {e}")
    
    def get_access_logs(self, limit: int = 100) -> list:
        """Recupera logs de acesso do sistema"""
        if not self.initialized:
            raise Exception("Sistema não foi inicializado")
        
        # Tentar Firebase primeiro se conectado
        if self.firebase_connected:
            try:
                ref = db.reference('access_logs')
                logs = ref.order_by_child('timestamp').limit_to_last(limit).get()
                
                if logs:
                    logs_list = []
                    for key, value in logs.items():
                        logs_list.append({
                            'id': key,
                            **value
                        })
                    
                    logs_list.sort(key=lambda x: x.get('timestamp', ''), reverse=True)
                    return logs_list
                    
            except Exception as e:
                print(f"⚠️ Firebase temporariamente indisponível: {e}")
                self.firebase_connected = False
        
        # Fallback para logs locais
        return self._get_local_logs(limit)
    
    def _get_local_logs(self, limit: int) -> list:
        """Recupera logs locais"""
        try:
            log_file = "local_access_log.json"
            
            if not os.path.exists(log_file):
                return []
            
            with open(log_file, 'r', encoding='utf-8') as f:
                logs = json.load(f)
            
            # Ordenar por timestamp e limitar
            logs.sort(key=lambda x: x.get('timestamp', ''), reverse=True)
            return logs[:limit]
            
        except Exception as e:
            print(f"Erro ao carregar logs locais: {e}")
            return []
    
    def get_user_access_stats(self, usuario: str) -> Dict[str, Any]:
        """Retorna estatísticas de acesso de um usuário"""
        if not self.initialized:
            raise Exception("Sistema não foi inicializado")
        
        # Tentar Firebase primeiro se conectado
        if self.firebase_connected:
            try:
                ref = db.reference('access_logs')
                user_logs = ref.order_by_child('usuario').equal_to(usuario).get()
                
                if user_logs:
                    total_acessos = len(user_logs)
                    timestamps = [log.get('timestamp', '') for log in user_logs.values()]
                    ips = list(set([log.get('ip', '') for log in user_logs.values()]))
                    
                    return {
                        'total_acessos': total_acessos,
                        'ultimo_acesso': max(timestamps) if timestamps else None,
                        'primeiro_acesso': min(timestamps) if timestamps else None,
                        'ips_utilizados': ips
                    }
                    
            except Exception as e:
                print(f"⚠️ Firebase temporariamente indisponível: {e}")
                self.firebase_connected = False
        
        # Fallback para logs locais
        return self._get_local_user_stats(usuario)
    
    def _get_local_user_stats(self, usuario: str) -> Dict[str, Any]:
        """Retorna estatísticas locais de um usuário"""
        try:
            logs = self._get_local_logs(1000)  # Buscar mais logs para stats
            user_logs = [log for log in logs if log.get('usuario') == usuario]
            
            if not user_logs:
                return {
                    'total_acessos': 0,
                    'ultimo_acesso': None,
                    'primeiro_acesso': None,
                    'ips_utilizados': []
                }
            
            total_acessos = len(user_logs)
            timestamps = [log.get('timestamp', '') for log in user_logs]
            ips = list(set([log.get('ip', '') for log in user_logs]))
            
            return {
                'total_acessos': total_acessos,
                'ultimo_acesso': max(timestamps) if timestamps else None,
                'primeiro_acesso': min(timestamps) if timestamps else None,
                'ips_utilizados': ips
            }
            
        except Exception as e:
            print(f"Erro ao calcular stats locais: {e}")
            return {
                'total_acessos': 0,
                'ultimo_acesso': None,
                'primeiro_acesso': None,
                'ips_utilizados': []
            }
    
    def sync_to_firebase(self):
        """Sincroniza logs locais com Firebase"""
        if not self.firebase_connected:
            print("⚠️ Firebase não conectado para sincronização")
            return
        
        try:
            local_logs = self._get_local_logs(1000)
            if not local_logs:
                return
            
            ref = db.reference('access_logs')
            for log in local_logs:
                ref.push(log)
            
            print(f"✅ {len(local_logs)} logs sincronizados com Firebase")
            
        except Exception as e:
            print(f"Erro na sincronização: {e}")
    
    def clear_all_logs(self):
        """Limpa todos os logs (local e Firebase)"""
        try:
            # Limpar logs locais
            self._clear_local_logs()
            
            # Limpar logs do Firebase se conectado
            if self.firebase_connected and self.db:
                try:
                    ref = self.db.reference('/access_logs')
                    ref.delete()
                    print("✅ Logs do Firebase limpos!")
                except Exception as e:
                    print(f"Erro ao limpar Firebase: {e}")
            
            print("✅ Todos os logs foram limpos!")
            return True
            
        except Exception as e:
            print(f"Erro ao limpar logs: {e}")
            return False

# Instância global do gerenciador Firebase
firebase_manager = FirebaseManager()