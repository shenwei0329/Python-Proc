
@echo off
start /wait "提示" cmd /c "mode con cols=80 lines=5 &echo. 《请确定！》 即将录入【%1】基础数据，按任意键继续&pause>nul"

python doXLSX2018.py d:\shenwei\2017-华云\研发管理数据\基础数据\%1【模板】.xlsx

