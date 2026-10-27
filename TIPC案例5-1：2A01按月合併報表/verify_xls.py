import importlib.util, pathlib, subprocess
spec=importlib.util.spec_from_file_location('app',pathlib.Path(__file__).with_name('server.py'))
app=importlib.util.module_from_spec(spec);spec.loader.exec_module(app)
src=next((pathlib.Path(__file__).parent.parent/'附件'/'CH5-1附件').glob('輸入_*.xlsx'))
profile=(pathlib.Path(app.TEMP.name)/'xls-test-profile').as_uri()
p=subprocess.run(['C:/Program Files/LibreOffice/program/soffice.exe','-env:UserInstallation='+profile,'--headless','--convert-to','xls:MS Excel 97','--outdir',app.TEMP.name,str(src)],capture_output=True,timeout=90)
path=pathlib.Path(app.TEMP.name)/(src.stem+'.xls')
assert path.exists(),(p.stdout,p.stderr)
f=app.read_file(path.name,path.read_bytes())
assert f['converted'] and f['sheets'][0]['month']==1
print('PASS: real BIFF xls converted and read with month and sheet intact.')
