# excel-read-write-util

## 介绍
常见的excel操作的方法封装，如解析合并单元格、解析excel中的图片等方法，都会写到此项目中，避免百度找来找去浪费时间精力。

## 目录说明
- 使用poi请在com.pangchun.poi目录下找。easy-excel的在com.pangchun.easyexcel目录下找。
- 两种excel工具都提供read、write的相关方法，请在对应目录下找。
- com.pangchun.poi / easyexcel.support目录下是一些支持类，如自定义异常、自定义注解等。
- 因为只是封装一些方法，因此没有连服务器，上传文件，excel模板都是本地模拟，使用resource目录下的static文件夹。
- 其他说明，根据文件夹命名很容易理解，此处不再作说明。


## 安装教程

- 导入相关maven或jar包，maven请参考pom.xml文件。
- 导入依赖后直接复制粘贴代码到自己项目中即可使用。
- 使用jdk1.8及以上。

## 使用说明

### poi -- 使用说明

#### read

- 通用的读方法请查询`package com.pangchun.poi.read.CommonRead<T>类`。

- 示例请查阅`package com.pangchun.poi.read.test包下测试类TestRead`。

  >- 此示例是将`resources/static/template/person-template.xlsx`此文件解析成Employee对象，图片上传至`resourcesstatic/image`下。
  >
  >- excel示例如图：
  >
  >  ![image-20210606175222009](assets/image-20210606175222009.png)
  >
  >- 测试类测试结果如下：
  >
  >  ```
  >  // 表头信息
  >  {0=员工信息导入模板}
  >  {0=编号, 1=姓名, 2=性别, 3=出生日期, 4=通讯地址, 5=联系方式, 6=所在部门, 7=基础薪资, 8=基础薪资抽取百分比, 9=基础薪资(抽取后), 10=证件照}
  >  
  >  // 正文信息
  >  Employee(id=sv3125, name=张远洋, sex=男, birth=1997-12-23 00:00:00, address=成都孵化园9座813号, phoneNumber=15282350478, departmentName=IT部, salary=20000, percent=0.22, salaryAfterPercent=15600.0, imageUrl=null)
  >  
  >  Employee(id=sv3145, name=欢嘉琦, sex=女, birth=1997-12-24 00:00:00, address=成都孵化园9座814号, phoneNumber=15282350478, departmentName=人事部, salary=9000, percent=0.20, salaryAfterPercent=7200.0, imageUrl=null)
  >  
  >  Employee(id=sv7145, name=范安喜, sex=女, birth=1997-12-25 00:00:00, address=成都孵化园9座815号, phoneNumber=15282350478, departmentName=财务部, salary=10000, percent=0.20, salaryAfterPercent=8000.0, imageUrl=null)
  >  
  >  // 图片信息
  >  ImageBean(firstRow=2, lastRow=3, firstCol=10, lastCol=11, url=F:\码云\excel-read-write-util\excel-util\src\main\resources\static\image\0.9287804194053441.jpeg)
  >  
  >  ImageBean(firstRow=2, lastRow=3, firstCol=10, lastCol=11, url=F:\码云\excel-read-write-util\excel-util\src\main\resources\static\image\0.4538432944414583.jpeg)
  >  
  >  ImageBean(firstRow=4, lastRow=4, firstCol=10, lastCol=10, url=F:\码云\excel-read-write-util\excel-util\src\main\resources\static\image\0.21050767499360856.jpeg)
  >  ```

#### write

### easyexcel -- 使用说明

#### read

#### write

## 参与贡献

- Fork 本仓库
- 新建 Feat_xxx 分支
- 提交代码
- 新建 Pull Request


## 其它

- 暂无。

