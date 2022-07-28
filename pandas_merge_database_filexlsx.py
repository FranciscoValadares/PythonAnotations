import pandas as pd
import pyodbc
import xlrd3 as xlrd


ServerName = '177.271.10.4'
MSQLDatabase = 'Name.Database'
username = 'name_usuario'
password = '*********'

## Connectar ao banco e recuparar cpf´s e e-mail
conn = pyodbc.connect("DRIVER={{SQL Server}};SERVER={0}; database={1}; \
                       trusted_connection=no;UID={2};PWD={3}".format(ServerName,MSQLDatabase,username,password))

df_usuarios = pd.read_sql_query('SELECT [cpf] ,[email] FROM [CFC.ACESSOS].[dbo].[usuario] group by cpf, email',conn)
print(df_usuarios.head())


## Todo reading of file in xlsx
my_sheet = 'Geral' 
file_name = 'D:\\Francisco\\desenv\\Python\\ProjetosCFC\\EPC\\RelatorioGeralEPCWEB_NaoCumpriu.xlsx' 
df_prof_naocumpriu = pd.read_excel(file_name, sheet_name = my_sheet)
df_prof_naocumpriu['cpf_only_number'] = df_prof_naocumpriu["CPF"].str.replace(".", "").str.replace("-", "")

##Save final result to excel
resultado_final = pd.merge(left=df_prof_naocumpriu, right=df_usuarios, left_on='cpf_only_number', right_on='cpf')
resultado_final = resultado_final.drop('cpf', axis=1)
resultado_final.to_excel(r'D:\\Francisco\\desenv\\Python\\ProjetosCFC\\EPC\\Rel_EPCWEB_NaoCumpriu_comEmail.xlsx', index = False)

 



#requirements.txt
#pyodbc==4.0.34
#pandas==1.4.3
#xlrd3==1.1.0
#openpyxl==3.0.10
