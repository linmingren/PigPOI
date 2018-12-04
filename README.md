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

下面的例子都用到了下面2个Java Bean

+ User
```Java
@Data
public class User {
    private String name;
    private String address;
    private int score;
    private Date createdAt;

    public User(String user, String address, int score, Date createdAt) {
        this.name = user;
        this.address = address;
        this.score = score;
        this.createdAt = createdAt;
    }
}
```
+ SalaryPayment
```Java
@Data
public class SalaryPayment {
    private String userName;
    private int baseSalary;
    private int fullAttendanceBonus;
    private int mealSupplement;
    private int transportationAllowance;
    private int sickLeave;
    private int personalLeave;

    public SalaryPayment(String userName, int baseSalary, int fullAttendanceBonus, int mealSupplement,
                         int transportationAllowance,  int sickLeave, int personalLeave) {
        this.userName = userName;
        this.baseSalary = baseSalary;
        this.fullAttendanceBonus = fullAttendanceBonus;
        this.mealSupplement = mealSupplement;
        this.transportationAllowance = transportationAllowance;
        this.sickLeave = sickLeave;
        this.personalLeave = personalLeave;
    }

    public int getActualPay() {
        return this.baseSalary + this.fullAttendanceBonus + this.mealSupplement + this.transportationAllowance
                - sickLeave - personalLeave;
    }
}
```

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

导出结果

![](https://raw.githubusercontent.com/linmingren/helloexcel/master/images/simpleTable.png)
## 自定义表头显示样式
```Java
        TableExcel excel = new TableExcel();
        TableSheet sheet = new TableSheet("sheet1");

        TableHeaderRow row = new TableHeaderRow(){
            @Override
            protected CellStyle updatedStyle() {
                HSSFCellStyle cellStyle = excel.getWorkbook().createCellStyle();
                cellStyle.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
                cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                return cellStyle;
            }
        };

        for (String t :Arrays.asList("姓名", "地址", "分数", "考试时间")) {
            row.addCell(new TableCell(t));
        }        


        sheet.addRow(row);

        List<User> userList = new ArrayList<>();
        userList.add(new User("老王", "隔壁", 59, new Date()));
        userList.add(new User("小明", "草地上", 80, new Date()));
        userList.add(new User("超人", "飞机上", 100, new Date()));

        sheet.setData(Arrays.asList("name", "address", "score", "createdAt"), userList);

        excel.addSheet(sheet);

        FileOutputStream output = new FileOutputStream("excels/customHeader.xls");
        excel.render(output);
        output.close();
```

导出结果

![](https://raw.githubusercontent.com/linmingren/helloexcel/master/images/customHeader.png)

## 合并表头

```Java
        TableExcel excel = new TableExcel();
        TableSheet sheet = new TableSheet("sheet1");

        TableRow row = new TableHeaderRow();
        TableCell cell = new TableCell("姓名");
        cell.setRowSpan(3);
        row.addCell(cell);
        cell = new TableCell("收入");
        cell.setColSpan(4);
        row.addCell(cell);
        cell = new TableCell("扣除");
        cell.setColSpan(2);
        row.addCell(cell);
        cell = new TableCell("实际发放");
        cell.setRowSpan(3);
        row.addCell(cell);
        sheet.addRow(row);

        row = new TableHeaderRow();

        cell = new TableCell("基本工资");
        cell.setRowSpan(2);
        row.addCell(cell);
        cell = new TableCell("补贴");
        cell.setColSpan(3);
        row.addCell(cell);
        cell = new TableCell("事假");
        cell.setRowSpan(2);
        row.addCell(cell);
        cell = new TableCell("病假");
        cell.setRowSpan(2);
        row.addCell(cell);
        sheet.addRow(row);

        row = new TableHeaderRow();
        cell = new TableCell("全勤");
        row.addCell(cell);

        cell = new TableCell("餐补");
        row.addCell(cell);
        cell = new TableCell("交通补助");
        row.addCell(cell);
        sheet.addRow(row);

        List<SalaryPayment> userList = new ArrayList<>();
        userList.add(new SalaryPayment("老王", 100000, 0, 1000, 200, 300, 400));
        userList.add(new SalaryPayment("小明", 10000, 1000, 1000, 300, 3000, 400));
        userList.add(new SalaryPayment("超人", 20000, 0, 1000, 0, 0, 400));


        sheet.setData(Arrays.asList("userName", "baseSalary", "fullAttendanceBonus", "mealSupplement", "transportationAllowance", "sickLeave", "personalLeave", "actualPay"),
                userList);


        excel.addSheet(sheet);

        FileOutputStream output = new FileOutputStream("excels/headerSpan.xls");
        excel.render(output);
        output.close();
```

导出结果

![](https://raw.githubusercontent.com/linmingren/helloexcel/master/images/headerSpan.png)

## 从模板到引入表头

## 自定义单元格样式

```Java
 private static class CustomTableSheet extends TableSheet {
        TableExcel workbook;
        public CustomTableSheet(String name, TableExcel workbook) {
            super(name);
            this.workbook = workbook;
        }

        //自定义数据行的显示效果
        @Override
        protected TableCell createDataCell(String fieldName, Object value, int row, int col) {
            //把优先级最高的效果放在最后面
            Map styleProperties = new HashMap();
            if (row % 2 == 0) {
                //偶数行背景是灰色
                styleProperties.put(FILL_FOREGROUND_COLOR,IndexedColors.GREY_25_PERCENT.getIndex());
                styleProperties.put(CellUtil.FILL_PATTERN, FillPatternType.SOLID_FOREGROUND);
            }

            if (col  == 0) {
                //第一列是粗体
                Font f = workbook.getWorkbook().createFont();
                f.setBold(true);
                styleProperties.put(FONT, f.getIndexAsInt());
            }

            if (fieldName.equals("score")) {
                if (Double.valueOf(value.toString()) < 60) {
                    //成绩少于60的背景是红色
                    styleProperties.put(FILL_FOREGROUND_COLOR,IndexedColors.RED.getIndex());
                    styleProperties.put(CellUtil.FILL_PATTERN, FillPatternType.SOLID_FOREGROUND);
                }
            }

            TableCell tableCell = new TableCell(value) {
                @Override
                public Map updatedStyle() {
                    return styleProperties;
                }
            };
            return tableCell;
        }
    }

    //如何自定义数据单元格显示样式的例子
    @Test
    public void customCell() throws IOException, InvocationTargetException, IllegalAccessException {
        TableExcel excel = new TableExcel();
        TableSheet sheet = new CustomTableSheet("sheet1", excel);

        TableRow row = TableHeaderRow.of(Arrays.asList("姓名", "地址", "分数", "考试时间"));
        sheet.addRow(row);

        List<User> userList = new ArrayList<>();
        userList.add(new User("老王", "隔壁", 59, new Date()));
        userList.add(new User("小明", "草地上", 80, new Date()));
        userList.add(new User("超人", "飞机上", 100, new Date()));
        userList.add(new User("蜘蛛侠", "飞机上", 50, new Date()));
        userList.add(new User("皇帝", "紫禁城", 100, new Date()));

        sheet.setData(Arrays.asList("name", "address", "score", "createdAt"), userList);

        excel.addSheet(sheet);

        FileOutputStream output = new FileOutputStream("excels/customCell.xls");
        excel.render(output);
        output.close();
    }
```

导出结果

![](https://raw.githubusercontent.com/linmingren/helloexcel/master/images/customCell.png)

## 性能
