# 解析coding导出的用例csv文件到word中
***
# 版本1 UseCaseExcelToWord_coding.py
## 所需环境：
- python环境:  
  python3 (需要联网：存在翻译模块)
  
- 需要的python库：pip install **

     1. python-docx    (生成word)
     2. csv   (读取用例导出excel)
     3. translate     (翻译中文到英文，用于生成英文标识符)
 ***
 
## 使用说明：
 
 1. 使用cmd打开命令行窗口
 2. 进入UseCaseExcelToWord_coding.py的目录，cd /UseCaseCsvToWord_coding
 3. 输入 python UseCaseExcelToWord_coding.py csv文件路径 word输出文件夹  
    例 python UseCaseExcelToWord_coding.py "coding完整用例导出.csv" "输出示例/"  
## 示例
![版本1导出的用例](https://github.com/1050035126/UseCaseCsvToWord_coding/blob/master/image/%E7%89%88%E6%9C%AC1.png)
***
# 版本2 UseCaseExcelToWord_coding2.py
## 所需环境：
- python环境:  
  python3
  
- 需要的python库：pip install **

     1. python-docx    (生成word)
     2. csv   (读取用例导出excel)
 ***
 
## 使用说明：
 
 1. 修改UseCaseExcelToWord_coding2.py文件中的变量
    1. signalPre 用例的前缀
    2. moduleSigDic 模块名称的英文缩写
    3. excelPath 从coding上导出的csv格式的测试用例位置
    4. wordOutDir 测试用例word生成文件夹路径
 2. 运行UseCaseExcelToWord_coding2.py
 
## 示例
![版本2导出的用例](https://github.com/1050035126/UseCaseCsvToWord_coding/blob/master/image/%E7%89%88%E6%9C%AC2.png)
 
