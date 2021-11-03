## 文件说明

### extract.py

这个文件用于提取文件中的信息，按照原始文件中的表头，修改 `headers`中的内容，运行该文件后从原始文件中提取`headers`中设置的对应列，请务必确保表头中严格包含`headers`中的关键字，**例如`姓名`和`名字`不等价。**

实现的功能

+ [x] 按照性别分开
+ [x] 简单纠错
  * 按照最新填报的信息覆盖相同学号的报名信息



### shuffle.py

这个文件用来打乱非种子选手，使用时请设置好`headers`中的内容，要求同上。建议配合上面的文件使用



## 运行环境

1. 请安装`python3`

2. 代码依赖`xlrd`和`xlwt`两个包，安装命令如下

   ```sh
   pip install xlrd==1.2.0 xlwt==1.3.0
   ```

   仅在上述版本中测试过，其他版本可能不支持`xlsx`格式文件，可以尝试将后缀改为`xls`。

运行

```sh
python file_name.py # 将 file_name.py 换成对应的文件名
```

