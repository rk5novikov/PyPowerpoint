import shutil
import win32com
from win32com import client

app = client.gencache.EnsureDispatch('Excel.Application')
constants = client.constants.__dict__['__dicts__'][0]



with open('xl_constants.py', 'w') as f0:
    for var, value in constants.items():
        if type(value) == int:
            f0.write('{0} = {1:d}\n'.format(var, value))

shutil.rmtree(win32com.__gen_path__)