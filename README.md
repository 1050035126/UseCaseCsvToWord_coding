# 解析coding导出的用例csv文件到word中
***
## 所需环境：
- python环境:  
  python3 (需要联网：存在翻译模块)
  
- 需要的python库：pip install **

     1. python-docx    (生成word)
     2. csv   (读取用例导出excel)
     3. translate     (翻译中文到英文，用于生成英文标识符)
 ***
 
 ##使用说明：
 
 1. 使用cmd打开命令行窗口
 2. 进入UseCaseExcelToWord_coding.py的目录，cd /UseCaseCsvToWord_coding
 3. 输入 python UseCaseExcelToWord_coding.py csv文件路径 word输出文件夹  
    例 python UseCaseExcelToWord_coding.py "coding完整用例导出.csv" "输出示例/"  
 
 