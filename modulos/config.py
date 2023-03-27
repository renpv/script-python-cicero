import os
from dotenv import load_dotenv
load_dotenv()  

# Definindo variáveis do script
base = os.environ.get('BASE_URL')
BASE_URL = base if base not in ['', None] else ''

# informe uma string em branco '' se quiser que o script solicite o login
USER_LOGIN = os.environ.get('USER_LOGIN')

# informe uma string em branco '' se quiser que o script solicite a senha
USER_PASS = os.environ.get('USER_PASS')

work_dir = os.path.dirname(os.path.abspath(__name__))