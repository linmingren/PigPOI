# PigPOI 
只要是开发过业务系统的同学，应该都开发过导出成Excel文件的功能（甚至你跳过的每家公司都开发了一次有没有？）。
这些功能看似不难，但是实现起来却有点复杂（因为大家使用的Apache POI是个抽象级别很低的Excel封装）。PigPOI的
出现就是为了让这其中80%的场景变得容易些。

废话不多说，目前支持的功能：

+ 能够定制表头的样式（对齐，字体，包括一些简单的合并单元格）
+ 能够把Java Bean的字段输出到对应的行
+ 能够对数据行的样式进行简单定制（比如根据的值的大小更改背景色等等）
+ 可以从模板文件中导入表头，这个功能对复杂表头设计特别有意义
+ 能够通过流来导出从而不需要太大的内存（你知道导出100w行数据需要多大内存吗？）

## 设计思想

+ 简单，不过度封装，同时应该能完成基本的功能

这个库[新一代 Excel 导出工具：ExcelUtil + RunnerUtil 介绍](https://juejin.im/post/5bfdf1aa6fb9a049a62c460f) 功能非常强大，但是明显是过度设计了，甚至用了简单的自定义脚本来实现。设想你用的过程中出现了问题，那个头真是大了。

[EasyPoi](https://gitee.com/lemur/easypoi) 也是类似的过度封装的库，一个导出Excel的库都需要加QQ群来问问题，想起来就头大。

相反，这个库[ExcelUtil](https://github.com/SargerasWang/ExcelUtil/blob/master/src/main/java/com/sargeraswang/util/ExcelUtil/ExcelUtil.java) 就显得太简单了，连最基本自定义样式和合并表头的功能都没有，你能
想象领导看到导出的Excel中标题和数据的格式是一样会是什么表情？

PigPOI的目的是解决业务中80%左右的导出问题，剩下的20% 就直接用POI乖乖自己写吧。

+ 无依赖

直接把代码嵌入到你的业务系统就能跑起来。


# 应用例子

这个工具能够支持的功能在例子中都写出来了，如果没有写出来的，就是不支持，不用去代码中找了，节省你的宝贵的时间。

## 最简单的导出

```Java
        TableExcel excel = new TableExcel();
        TableSheet sheet = new TableSheet("sheet1");

        TableRow row = TableHeaderRow.of(Arrays.asList("姓名", "地址", "分数", "考试时间"));
        sheet.addRow(row);

        List<User> userList = new ArrayList<>();
        userList.add(new User("老王", "隔壁", 59, new Date()));
        userList.add(new User("小明", "草地上", 80, new Date()));
        userList.add(new User("超人", "飞机上", 100, new Date()));

        sheet.setData(Arrays.asList("name", "address", "score", "createdAt"), userList);

        excel.addSheet(sheet);

        FileOutputStream output = new FileOutputStream("excels/simpleRender.xls");
        excel.render(output);
        output.close();
```

结果
![](https://raw.githubusercontent.com/linmingren/helloexcel/master/images/simpleTable.png)
## 自定义单元格的显示样式

## 合并表头

## 从模板到引入表头

## 性能
