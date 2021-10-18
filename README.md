# xmindToExcel

把xmind文件解析并写入到excel，主要用于把脑图的用例转换成excel格式的用例

依赖库：
xmindparser
xlwt

命令安装依赖
pip3 install xmindparser
pip3 install xlwt

运行：
1、先把需要转义的xmind文件放到files目录下
2、在parse.py中把XmindToExcel('需要转义的脑图文件路径')
3、转换后的文件在工程下的test.xls中

注意脑图需要参照files/模版.xmind这个格式来写哦。
