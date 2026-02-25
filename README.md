# 📊 Painel SGE - Sistema de Gestão Escolar

Um painel interativo desenvolvido em Streamlit para análise de notas, frequência e alertas escolares baseado em dados do Sistema de Gestão Escolar (SGE).

## 🚀 Acesso Online

**Link do Streamlit**: [Clique aqui para acessar o painel online](https://seu-usuario-painel-sge.streamlit.app/)

## 📋 Funcionalidades

- **📈 Análise de Notas**: Visualização de notas dos 1º e 2º bimestres
- **🚨 Alertas Críticos**: Identificação de alunos em risco de reprovação
- **📊 Análise de Frequência**: Monitoramento de frequência escolar
- **🔍 Filtros Avançados**: Por escola, turma, disciplina e aluno
- **📉 Gráficos Interativos**: Visualizações com Plotly
- **⚖️ Corda Bamba**: Cálculo de notas necessárias para aprovação

## 🛠️ Como Usar

### 1. Upload de Dados
- Faça upload de uma planilha Excel (.xlsx) com os dados do SGE
- Ou salve o arquivo como `dados.xlsx` na pasta do projeto

### 2. Estrutura da Planilha
A planilha deve conter as seguintes colunas:
- **Escola**: Nome da escola
- **Turma**: Nome da turma
- **Turno**: Turno de estudo
- **Aluno**: Nome do aluno
- **Período**: Bimestre (ex: "Primeiro Bimestre", "Segundo Bimestre")
- **Disciplina**: Nome da disciplina
- **Nota**: Nota do aluno (0-10)
- **Falta**: Número de faltas
- **Frequência**: Percentual de frequência
- **Status**: Status do aluno

### 3. Filtros
Use a barra lateral para filtrar por:
- Escola específica
- Status do aluno
- Turmas selecionadas
- Disciplinas específicas
- Aluno individual

## 📊 Indicadores

### Classificações de Notas
- **🟢 Verde**: Aluno aprovado (N1≥6 e N2≥6)
- **🔴 Vermelho Duplo**: Risco alto (N1<6 e N2<6)
- **🟡 Queda p/ Vermelho**: Piorou (N1≥6 e N2<6)
- **🔵 Recuperou**: Melhorou (N1<6 e N2≥6)
- **⚪ Incompleto**: Falta nota

### Classificações de Frequência
- **🔴 < 75%**: Reprovado por frequência
- **🟠 < 80%**: Alto risco de reprovação
- **🟡 < 90%**: Risco moderado
- **🟠 < 95%**: Ponto de atenção
- **🟢 ≥ 95%**: Meta favorável

## 🚀 Deploy Local

### Pré-requisitos
- Python 3.8 ou superior
- pip (gerenciador de pacotes Python)

### Instalação
```bash
# Clone o repositório
git clone https://github.com/seu-usuario/painel-sge.git
cd painel-sge

# Instale as dependências
pip install -r requirements.txt

# Execute o painel
streamlit run app.py
```

### Acesso Local
Abra seu navegador em: `http://localhost:8501`

## 📦 Dependências

- **pandas**: Manipulação de dados
- **streamlit**: Framework web
- **openpyxl**: Leitura de arquivos Excel
- **plotly**: Gráficos interativos
- **numpy**: Operações numéricas

## 🔧 Configurações

### Médias de Aprovação
```python
MEDIA_APROVACAO = 6.0  # Média para aprovação
MEDIA_FINAL_ALVO = 6.0  # Média final desejada
```

### Personalização
Você pode ajustar as constantes no início do arquivo `app.py` para:
- Alterar a média de aprovação
- Modificar critérios de frequência
- Ajustar cores e estilos

## 📱 Responsividade

O painel é totalmente responsivo e funciona em:
- 💻 Desktop
- 📱 Tablets
- 📱 Smartphones

## 🤝 Contribuição

1. Faça um fork do projeto
2. Crie uma branch para sua feature (`git checkout -b feature/AmazingFeature`)
3. Commit suas mudanças (`git commit -m 'Add some AmazingFeature'`)
4. Push para a branch (`git push origin feature/AmazingFeature`)
5. Abra um Pull Request

## 📄 Licença

Este projeto está sob a licença MIT. Veja o arquivo `LICENSE` para mais detalhes.

## 👨‍💻 Desenvolvedor

**Alexandre Tolentino**
- Desenvolvido para facilitar a análise de dados escolares
- Sistema de Gestão Escolar (SGE)

## 📞 Suporte

Se encontrar algum problema ou tiver sugestões:
1. Abra uma [Issue](https://github.com/seu-usuario/painel-sge/issues)
2. Entre em contato via email
3. Consulte a documentação do Streamlit

---

⭐ **Se este projeto foi útil, considere dar uma estrela no GitHub!**

