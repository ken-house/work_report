# work_report
从TAPD导出多个Excel文件，最终生成工作报表；
## 使用说明
执行命令生成可执行文件tapd.exe
```shell
go build -o tapd.exe
```
执行tapd.exe文件，会读取当前目录下的input_file目录下的文件
```shell
./tapd.exe
```
这样会生成一个VR研发中心每日工作.xlsx文件。

若要指定Excel文件目录及生成文件路径，可以指定-dir_path、-save_path参数；
```shell
./tapd.exe  -dir_path E:\work\tapd\input_file  -save_path  E:\测试.xlsx
```