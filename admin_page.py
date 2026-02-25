"""
Página Admin - Monitoramento de Acessos
"""
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
from io import BytesIO
import json
from firebase_config import firebase_manager
from ip_utils import get_client_info

def tela_admin():
    """Tela de login para administradores"""
    st.markdown("""
    <div style="text-align: center; padding: 40px 20px; background: linear-gradient(135deg, #dc2626, #ef4444); border-radius: 15px; margin-bottom: 30px;">
        <h1 style="color: white; margin: 0; font-size: 2.5em; font-weight: 700;">🔐 Painel Administrativo</h1>
        <h2 style="color: white; margin: 15px 0 0 0; font-weight: 600;">Monitoramento de Acessos</h2>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        st.markdown("### Acesso Administrativo")
        st.warning("⚠️ Esta área é restrita apenas para administradores")
        
        with st.form("admin_login_form"):
            admin_user = st.text_input("Usuário Admin:", placeholder="admin")
            admin_password = st.text_input("Senha Admin:", type="password", placeholder="Digite a senha administrativa")
            
            col_btn1, col_btn2 = st.columns(2)
            with col_btn1:
                login_btn = st.form_submit_button("Entrar como Admin", use_container_width=True, type="primary")
            with col_btn2:
                if st.form_submit_button("Voltar", use_container_width=True):
                    st.session_state.admin_logado = False
                    st.session_state.mostrar_admin = False
                    st.rerun()
        
        if login_btn:
            # Verificação simples de admin (você pode melhorar isso)
            if admin_user == "admin" and admin_password == "admin123":
                st.session_state.admin_logado = True
                st.success("Login administrativo realizado com sucesso!")
                st.rerun()
            else:
                st.error("Usuário ou senha administrativa incorretos!")

def dashboard_admin():
    """Dashboard principal do administrador"""
    st.markdown("""
    <div style="text-align: center; padding: 30px 20px; background: linear-gradient(135deg, #dc2626, #ef4444); border-radius: 15px; margin-bottom: 30px;">
        <h1 style="color: white; margin: 0; font-size: 2.2em; font-weight: 700;">📊 Dashboard Administrativo</h1>
        <h2 style="color: white; margin: 10px 0 0 0; font-weight: 600;">Monitoramento de Acessos em Tempo Real</h2>
    </div>
    """, unsafe_allow_html=True)
    
    # Botões de controle
    col_control1, col_control2 = st.columns([3, 1])
    
    with col_control1:
        if st.button("👥 Estatísticas por Usuário", use_container_width=True, type="primary"):
            st.session_state.mostrar_stats_usuario = True
            st.rerun()
    
    with col_control2:
        if st.button("🚪 Sair do Admin", use_container_width=True):
            st.session_state.admin_logado = False
            st.session_state.mostrar_admin = False
            st.rerun()
    
    st.markdown("---")
    
    try:
        # Carregar dados do Firebase
        with st.spinner("Carregando dados de monitoramento..."):
            logs = firebase_manager.get_access_logs(limit=500)
        
        if not logs:
            st.warning("Nenhum log de acesso encontrado ainda.")
            return
        
        # Converter para DataFrame
        df_logs = pd.DataFrame(logs)
        
        # Converter timestamp de forma simples
        df_logs['timestamp'] = pd.to_datetime(df_logs['timestamp'], errors='coerce')
        df_logs = df_logs.dropna(subset=['timestamp'])
        
        if len(df_logs) == 0:
            st.warning("Nenhum timestamp válido encontrado nos logs.")
            return
            
        df_logs['data'] = df_logs['timestamp'].dt.date
        df_logs['hora'] = df_logs['timestamp'].dt.strftime('%H:%M')
        
        # Métricas principais
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            total_acessos = len(df_logs)
            st.metric("Total de Acessos", total_acessos)
        
        with col2:
            usuarios_unicos = df_logs['usuario'].nunique()
            st.metric("Usuários Únicos", usuarios_unicos)
        
        with col3:
            ips_unicos = df_logs['ip'].nunique()
            st.metric("IPs Únicos", ips_unicos)
        
        with col4:
            hoje = datetime.now().date()
            acessos_hoje = len(df_logs[df_logs['data'] == hoje])
            st.metric("Acessos Hoje", acessos_hoje)
        
        st.markdown("---")
        
        # Filtros
        col_filter1, col_filter2, col_filter3 = st.columns(3)
        
        with col_filter1:
            usuarios_disponiveis = ['Todos'] + sorted(df_logs['usuario'].unique().tolist())
            usuario_filtro = st.selectbox("Filtrar por Usuário:", usuarios_disponiveis)
        
        with col_filter2:
            datas_disponiveis = sorted(df_logs['data'].unique(), reverse=True)
            data_filtro = st.selectbox("Filtrar por Data:", ['Todas'] + [str(d) for d in datas_disponiveis])
        
        with col_filter3:
            ips_disponiveis = ['Todos'] + sorted(df_logs['ip'].unique().tolist())
            ip_filtro = st.selectbox("Filtrar por IP:", ips_disponiveis)
        
        # Aplicar filtros
        df_filtrado = df_logs.copy()
        
        if usuario_filtro != 'Todos':
            df_filtrado = df_filtrado[df_filtrado['usuario'] == usuario_filtro]
        
        if data_filtro != 'Todas':
            data_selecionada = pd.to_datetime(data_filtro).date()
            df_filtrado = df_filtrado[df_filtrado['data'] == data_selecionada]
        
        if ip_filtro != 'Todos':
            df_filtrado = df_filtrado[df_filtrado['ip'] == ip_filtro]
        
        # Gráficos
        col_graph1, col_graph2 = st.columns(2)
        
        with col_graph1:
            # Gráfico de acessos por dia
            acessos_por_dia = df_filtrado.groupby('data').size().reset_index(name='acessos')
            fig_dia = px.line(acessos_por_dia, x='data', y='acessos', 
                             title='Acessos por Dia', markers=True)
            fig_dia.update_layout(xaxis_title="Data", yaxis_title="Número de Acessos")
            st.plotly_chart(fig_dia, use_container_width=True)
        
        with col_graph2:
            # Gráfico de acessos por usuário
            acessos_por_usuario = df_filtrado.groupby('usuario').size().reset_index(name='acessos')
            fig_usuario = px.bar(acessos_por_usuario, x='usuario', y='acessos',
                                title='Acessos por Usuário')
            fig_usuario.update_layout(xaxis_title="Usuário", yaxis_title="Número de Acessos")
            fig_usuario.update_xaxis(tickangle=45)
            st.plotly_chart(fig_usuario, use_container_width=True)
        
        # Gráfico de acessos por hora
        df_filtrado['hora_int'] = df_filtrado['timestamp'].dt.hour
        acessos_por_hora = df_filtrado.groupby('hora_int').size().reset_index(name='acessos')
        fig_hora = px.bar(acessos_por_hora, x='hora_int', y='acessos',
                         title='Acessos por Hora do Dia')
        fig_hora.update_layout(xaxis_title="Hora", yaxis_title="Número de Acessos")
        st.plotly_chart(fig_hora, use_container_width=True)
        
        st.markdown("---")
        
        # Tabela de logs recentes
        st.markdown("### 📋 Logs de Acesso Recentes")
        
        # Preparar dados para exibição
        df_exibicao = df_filtrado[['data_hora', 'usuario', 'ip', 'user_agent']].copy()
        df_exibicao.columns = ['Data/Hora', 'Usuário', 'IP', 'Navegador']
        df_exibicao = df_exibicao.sort_values('Data/Hora', ascending=False)
        
        st.dataframe(df_exibicao, use_container_width=True, height=400)
        
        # Botões de ação
        col_export, col_clean = st.columns(2)
        
        with col_export:
            if st.button("📥 Exportar Logs para Excel"):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_exibicao.to_excel(writer, sheet_name='Logs de Acesso', index=False)
                
                st.download_button(
                    label="⬇️ Baixar Arquivo Excel",
                    data=output.getvalue(),
                    file_name=f"logs_acesso_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        
        with col_clean:
            if st.button("🧹 Limpar Logs Duplicados"):
                try:
                    # Limpar logs duplicados (manter apenas um por usuário a cada 2 minutos)
                    logs_limpos = []
                    for log in logs:
                        usuario = log.get('usuario', '')
                        timestamp = log.get('timestamp', '')
                        
                        # Verificar se já existe um log similar recente
                        log_similar = False
                        for log_existente in logs_limpos:
                            if log_existente.get('usuario') == usuario:
                                try:
                                    # Usar a mesma função de parsing de timestamp
                                    ts1 = parse_timestamp(timestamp)
                                    ts2 = parse_timestamp(log_existente.get('timestamp', ''))
                                    if abs((ts1 - ts2).seconds) < 120:
                                        log_similar = True
                                        break
                                except:
                                    # Se não conseguir comparar timestamps, considerar como similar
                                    log_similar = True
                                    break
                        
                        if not log_similar:
                            logs_limpos.append(log)
                    
                    # Salvar logs limpos
                    with open('local_access_log.json', 'w', encoding='utf-8') as f:
                        json.dump(logs_limpos, f, ensure_ascii=False, indent=2)
                    
                    st.success(f"Logs limpos! Removidos {len(logs) - len(logs_limpos)} duplicados.")
                    st.rerun()
                    
                except Exception as e:
                    st.error(f"Erro ao limpar logs: {e}")
        
        # Botão para sincronizar com Firebase
        col_sync, col_empty = st.columns(2)
        
        with col_sync:
            if st.button("☁️ Sincronizar com Firebase"):
                try:
                    firebase_manager.sync_to_firebase()
                    st.success("✅ Dados sincronizados com Firebase!")
                    st.info("Atualize a página do Firebase Console para ver os dados.")
                except Exception as e:
                    st.error(f"Erro na sincronização: {e}")
                    st.info("Os dados continuam salvos localmente no arquivo 'local_access_log.json'")
    
    except Exception as e:
        st.error(f"Erro ao carregar dados: {str(e)}")
        st.info("Verifique se o Firebase está configurado corretamente.")

def relatorio_completo():
    """Relatório completo de acessos"""
    st.markdown("### 📊 Relatório Completo de Acessos")
    
    # Botões de ação no topo
    col_btn1, col_btn2, col_btn3 = st.columns(3)
    
    with col_btn1:
        if st.button("🔄 Atualizar Relatório", use_container_width=True, type="primary"):
            st.rerun()
    
    with col_btn2:
        if st.button("🗑️ RESETAR DADOS (ZERAR TUDO)", use_container_width=True, type="secondary"):
            if st.session_state.get('confirm_reset', False):
                # Confirmar reset
                try:
                    # Limpar arquivo local
                    with open('local_access_log.json', 'w', encoding='utf-8') as f:
                        json.dump([], f, ensure_ascii=False, indent=2)
                    
                    # Tentar limpar Firebase também
                    try:
                        firebase_manager.clear_all_logs()
                    except:
                        pass
                    
                    st.success("✅ Dados resetados com sucesso!")
                    st.session_state.confirm_reset = False
                    st.rerun()
                except Exception as e:
                    st.error(f"Erro ao resetar: {e}")
            else:
                st.session_state.confirm_reset = True
                st.warning("⚠️ Clique novamente para confirmar o RESET!")
    
    with col_btn3:
        if st.button("← Voltar ao Dashboard", use_container_width=True, key="btn_voltar_relatorio_admin"):
            st.session_state.mostrar_relatorio = False
            st.rerun()
    
    try:
        logs = firebase_manager.get_access_logs(limit=1000)
        
        if not logs:
            st.warning("Nenhum log encontrado.")
            return
        
        # Converter para DataFrame simples
        df_logs = pd.DataFrame(logs)
        
        # Criar lista de todos os acessos (como planilha)
        st.markdown("#### 📋 Lista Completa de Todos os Acessos")
        
        # Preparar dados para exibição
        df_display = df_logs[['data_hora', 'usuario', 'ip', 'user_agent']].copy()
        df_display.columns = ['Data/Hora', 'Usuário', 'IP', 'Navegador']
        df_display = df_display.sort_values('Data/Hora', ascending=False)
        
        # Exibir tabela completa
        st.dataframe(
            df_display, 
            use_container_width=True, 
            height=500,
            hide_index=True
        )
        
        # Estatísticas resumidas
        col_stats1, col_stats2, col_stats3, col_stats4 = st.columns(4)
        
        with col_stats1:
            total_acessos = len(df_logs)
            st.metric("Total de Acessos", total_acessos)
        
        with col_stats2:
            usuarios_unicos = df_logs['usuario'].nunique()
            st.metric("Usuários Únicos", usuarios_unicos)
        
        with col_stats3:
            ips_unicos = df_logs['ip'].nunique()
            st.metric("IPs Únicos", ips_unicos)
        
        with col_stats4:
            if len(df_logs) > 0:
                ultimo_acesso = df_logs['data_hora'].iloc[0]  # Primeiro da lista ordenada
                st.metric("Último Acesso", ultimo_acesso)
        
        # Botões de exportação
        col_export1, col_export2 = st.columns(2)
        
        with col_export1:
            if st.button("📥 Exportar Lista Completa para Excel"):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_display.to_excel(writer, sheet_name='Todos os Acessos', index=False)
                
                st.download_button(
                    label="⬇️ Baixar Lista Completa",
                    data=output.getvalue(),
                    file_name=f"todos_acessos_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        
        with col_export2:
            # Lista resumida por usuário
            stats_usuarios = df_logs.groupby('usuario').agg({
                'data_hora': ['count', 'first', 'last']
            })
            stats_usuarios.columns = ['Total_Acessos', 'Primeiro_Acesso', 'Ultimo_Acesso']
            stats_usuarios = stats_usuarios.reset_index()
            stats_usuarios.columns = ['Usuário', 'Total de Acessos', 'Primeiro Acesso', 'Último Acesso']
            stats_usuarios = stats_usuarios.sort_values('Total de Acessos', ascending=False)
            
            if st.button("📊 Exportar Resumo por Usuário"):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    stats_usuarios.to_excel(writer, sheet_name='Resumo por Usuário', index=False)
                
                st.download_button(
                    label="⬇️ Baixar Resumo",
                    data=output.getvalue(),
                    file_name=f"resumo_usuarios_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        
    except Exception as e:
        st.error(f"Erro ao gerar relatório: {str(e)}")

def estatisticas_usuario():
    """Estatísticas detalhadas por usuário"""
    st.markdown("### 👥 Estatísticas por Usuário")
    
    # Botão para voltar ao dashboard
    if st.button("⬅️ Voltar ao Dashboard", key="btn_voltar_stats_usuario_admin"):
        st.session_state.mostrar_stats_usuario = False
        st.rerun()
    
    try:
        logs = firebase_manager.get_access_logs(limit=1000)
        
        if not logs:
            st.warning("Nenhum log encontrado.")
            return
        
        df_logs = pd.DataFrame(logs)
        
        # Lista de usuários únicos
        usuarios_unicos = sorted(df_logs['usuario'].unique())
        
        # Campo de busca por nome
        st.markdown("#### 🔍 Buscar Usuário")
        busca_nome = st.text_input("Digite o nome para buscar:", placeholder="Ex: ALEXANDRE")
        
        # Filtrar usuários baseado na busca
        if busca_nome:
            usuarios_filtrados = [u for u in usuarios_unicos if busca_nome.upper() in u.upper()]
        else:
            usuarios_filtrados = usuarios_unicos
        
        if not usuarios_filtrados:
            st.warning("Nenhum usuário encontrado com esse nome.")
            return
        
        # Selecionar usuário
        usuario_selecionado = st.selectbox("Selecionar usuário:", usuarios_filtrados)
        
        if usuario_selecionado:
            # Estatísticas do usuário
            stats = firebase_manager.get_user_access_stats(usuario_selecionado)
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.metric("Total de Acessos", stats['total_acessos'])
            
            with col2:
                if stats['ultimo_acesso']:
                    ultimo_acesso = pd.to_datetime(stats['ultimo_acesso'])
                    st.metric("Último Acesso", ultimo_acesso.strftime('%d/%m/%Y %H:%M'))
                else:
                    st.metric("Último Acesso", "N/A")
            
            with col3:
                st.metric("IPs Utilizados", len(stats['ips_utilizados']))
            
            # IPs utilizados
            st.markdown("#### 🌐 IPs Utilizados")
            for ip in stats['ips_utilizados']:
                st.write(f"• {ip}")
            
            # Histórico do usuário
            st.markdown("#### 📋 Histórico de Acessos")
            
            df_usuario = df_logs[df_logs['usuario'] == usuario_selecionado].copy()
            df_usuario = df_usuario.sort_values('timestamp', ascending=False)
            
            df_exibicao = df_usuario[['data_hora', 'ip', 'user_agent']].copy()
            df_exibicao.columns = ['Data/Hora', 'IP', 'Navegador']
            
            st.dataframe(df_exibicao, use_container_width=True)
    
    except Exception as e:
        st.error(f"Erro ao carregar estatísticas: {str(e)}")
