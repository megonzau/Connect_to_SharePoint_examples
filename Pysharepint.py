from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
import io
import pandas as pd


url = 'https://eafit.sharepoint.com/sites/Proyectoinformedecoyunturaeconomica'
username = 'megonzalea@eafit.edu.co'
password = ''
relative_url = r'/sites/Proyectoinformedecoyunturaeconomica/Documentos compartidos/PowerBI/Data/Base de datos.xlsx'


ctx_auth = AuthenticationContext(url)
ctx_auth.acquire_token_for_user(username, password)
ctx = ClientContext(url, ctx_auth)
web = ctx.web
ctx.load(web)
ctx.execute_query()


response = File.open_binary(ctx, relative_url)
#save data to BytesIO stream
bytes_file_obj = io.BytesIO()
bytes_file_obj.write(response.content)
bytes_file_obj.seek(0) #set file object to start

df_embig = pd.read_excel(bytes_file_obj, sheet_name = '25.EMBIG') 


df_ise = pd.read_excel(bytes_file_obj, sheet_name = '29.Indice de calidad y cu') 
df_ise["date"] = pd.period_range('2016-01-01', '2020-12-01', freq='M')



