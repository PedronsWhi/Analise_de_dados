import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.cluster import KMeans

# Carregar o arquivo Excel
file_path = 'Estágio CRM.xlsx'
df_clientes = pd.read_excel(file_path, sheet_name='Clientes')
df_interacoes = pd.read_excel(file_path, sheet_name='Interações')
df_compras = pd.read_excel(file_path, sheet_name='Compras')
df_marketing = pd.read_excel(file_path, sheet_name='Campanhas de Marketing')
df_feedback = pd.read_excel(file_path, sheet_name='Feedback dos Clientes')

# Exportar DataFrames para arquivos CSV
df_clientes.to_csv('clientes.csv', index=False)
df_interacoes.to_csv('interacoes.csv', index=False)
df_compras.to_csv('compras.csv', index=False)
df_marketing.to_csv('marketing.csv', index=False)
df_feedback.to_csv('feedback.csv', index=False)

# Verificar as primeiras linhas de cada DataFrame
print(df_clientes.head())
print(df_interacoes.head())
print(df_compras.head())
print(df_marketing.head())
print(df_feedback.head())

# Informações gerais
print(df_clientes.info())
print(df_interacoes.info())
print(df_compras.info())
print(df_marketing.info())
print(df_feedback.info())

# Estatísticas descritivas
print(df_clientes.describe())
print(df_interacoes.describe())
print(df_compras.describe())
print(df_marketing.describe())
print(df_feedback.describe())

# Visualização dos dados
sns.countplot(x='Gênero', data=df_clientes)
plt.show()

sns.histplot(df_clientes['Idade'], bins=10)
plt.show()

sns.histplot(df_clientes['Renda'], bins=10)
plt.show()

# Segmentação de Clientes
X = df_clientes[['Idade', 'Renda']]
kmeans = KMeans(n_clusters=3, random_state=0).fit(X)
df_clientes['Segmento'] = kmeans.labels_

# Visualizar os clusters
plt.scatter(df_clientes['Idade'], df_clientes['Renda'], c=df_clientes['Segmento'], cmap='viridis')
plt.xlabel('Idade')
plt.ylabel('Renda')
plt.title('Segmentação de Clientes')
plt.show()

# Avaliação das Campanhas de Marketing
# Converter a coluna 'Custo' para float
df_marketing['Custo'] = df_marketing['Custo'].replace(r'[\R$,]', '', regex=True).astype(float)
marketing_eff = df_marketing.groupby('Canal de Marketing')['Resultado (Taxa de Conversão)'].mean()
print(marketing_eff)

# Visualização do ROI
marketing_eff.plot(kind='bar')
plt.ylabel('ROI Médio')
plt.title('Eficácia dos Canais de Marketing')
plt.show()

# Salvar os resultados em um novo arquivo Excel
df_clientes.to_excel('Clientes_Segmentados.xlsx')
