import streamlit as st

st.set_page_config(page_title="Teste - Painel SGE", layout="wide")

st.title("🧪 Teste do Painel SGE")
st.success("✅ Streamlit está funcionando!")
st.info("Se você está vendo esta mensagem, o deploy está correto!")

st.markdown("""
## 📋 Próximos Passos:
1. ✅ Verificar se os arquivos estão no GitHub
2. ✅ Mudar branch de 'master' para 'main'
3. ✅ Fazer deploy no Streamlit
4. ✅ Testar o painel principal
""")

st.balloons()

