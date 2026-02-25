import pandas as pd

# Ler a planilha
fp = r"C:\Users\alexa\OneDrive\Área de Trabalho\Painel notas - Copia\AtaMapa (24).xlsx"
df = pd.read_excel(fp)

print("=" * 60)
print("ANÁLISE DO 3º BIMESTRE")
print("=" * 60)

# Filtrar apenas o 3º bimestre
terceiro_bim = df[df['Periodo'].str.contains('Terceiro', case=False, na=False)]

print(f"\n📊 Total de registros no 3º Bimestre: {len(terceiro_bim)}")

# Contar notas abaixo de 6
abaixo_6 = terceiro_bim[terceiro_bim['Nota'] < 6]
acima_igual_6 = terceiro_bim[terceiro_bim['Nota'] >= 6]

print(f"\n🔴 Notas ABAIXO de 6.0: {len(abaixo_6)} ({len(abaixo_6)/len(terceiro_bim)*100:.1f}%)")
print(f"🟢 Notas ACIMA ou IGUAL a 6.0: {len(acima_igual_6)} ({len(acima_igual_6)/len(terceiro_bim)*100:.1f}%)")

print(f"\n📈 Estatísticas das notas no 3º Bimestre:")
print(f"   Média: {terceiro_bim['Nota'].mean():.2f}")
print(f"   Mediana: {terceiro_bim['Nota'].median():.2f}")
print(f"   Mínima: {terceiro_bim['Nota'].min():.2f}")
print(f"   Máxima: {terceiro_bim['Nota'].max():.2f}")

print(f"\n✅ RESULTADO: No 3º bimestre tem MAIS registros {'ABAIXO' if len(abaixo_6) > len(acima_igual_6) else 'ACIMA'} da média!")
