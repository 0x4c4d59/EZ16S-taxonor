```bash
# 生成所有使用包名称txt文件
pip freeze > requirements.txt
```
```bash
# 打包文件
pyinstaller --onefile --windowed main.py
```