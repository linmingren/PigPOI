package me.linmingren.table.example;


import me.linmingren.table.*;
import me.linmingren.table.example.model.SalaryPayment;
import me.linmingren.table.example.model.User;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellUtil;
import org.junit.Test;

import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.util.*;

import static org.apache.poi.ss.util.CellUtil.FILL_FOREGROUND_COLOR;

/**
 * Unit test for simple App.
 */
public class TableExcelTest {

    //最简单的例子, 默认表头样式
    @Test
    public void simpleRender() throws IOException, InvocationTargetException, IllegalAccessException {
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
    }

    //自定义表头样式
    @Test
    public void customHeader() throws IOException, InvocationTargetException, IllegalAccessException {
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

        FileOutputStream output = new FileOutputStream("e:/test1.xls");
        excel.render(output);
        output.close();
    }

    private static class CustomTableSheet extends TableSheet {
        public CustomTableSheet(String name) {
            super(name);
        }

        //自定义数据行的显示效果
        @Override
        protected TableCell createDataCell(String fieldName, Object value, int row, int col) {
            if (fieldName.equals("score")) {
                TableCell tableCell = new TableCell(value) {
                    @Override
                    public Map updatedStyle() {
                        Map styleProperties = new HashMap();
                        if (Double.valueOf(value.toString()) < 60) {
                            //成绩少于60的背景是红色
                            styleProperties.put(FILL_FOREGROUND_COLOR,IndexedColors.RED.getIndex());
                            styleProperties.put(CellUtil.FILL_PATTERN, FillPatternType.SOLID_FOREGROUND);
                        }

                        return styleProperties;
                    }
                };
                return tableCell;
            } else if (col == 0) {
                //第一列是粗体
                TableCell tableCell = new TableCell(value) {
                    @Override
                    public Map updatedStyle() {
                        Map styleProperties = new HashMap();
                        if (Double.valueOf(value.toString()) < 60) {
                            //成绩少于60的背景是红色
                            styleProperties.put(FILL_FOREGROUND_COLOR,IndexedColors.RED.getIndex());
                            styleProperties.put(CellUtil.FILL_PATTERN, FillPatternType.SOLID_FOREGROUND);
                        }

                        return styleProperties;
                    }
                };
                return tableCell;
            } else if (row %2 == 0) {
                //偶数行背景是灰色
                TableCell tableCell = new TableCell(value) {
                    @Override
                    public Map updatedStyle() {
                        Map styleProperties = new HashMap();
                        if (Double.valueOf(value.toString()) < 60) {
                            //成绩少于60的背景是红色
                            styleProperties.put(FILL_FOREGROUND_COLOR,IndexedColors.RED.getIndex());
                            styleProperties.put(CellUtil.FILL_PATTERN, FillPatternType.SOLID_FOREGROUND);
                        }

                        return styleProperties;
                    }
                };
                return tableCell;
            }else {
                return super.createDataCell(fieldName, value, row, col);
            }
        }
    }

    //如何自定义数据单元格显示样式的例子
    @Test
    public void simpleRenderWithCustomCell() throws IOException, InvocationTargetException, IllegalAccessException {
        TableExcel excel = new TableExcel();
        TableSheet sheet = new CustomTableSheet("sheet1");

        TableRow row = TableHeaderRow.of(Arrays.asList("姓名", "地址", "分数", "考试时间"));
        sheet.addRow(row);

        List<User> userList = new ArrayList<>();
        userList.add(new User("老王", "隔壁", 59, new Date()));
        userList.add(new User("小明", "草地上", 80, new Date()));
        userList.add(new User("超人", "飞机上", 100, new Date()));

        sheet.setData(Arrays.asList("name", "address", "score", "createdAt"), userList);

        excel.addSheet(sheet);

        FileOutputStream output = new FileOutputStream("e:/test3.xls");
        excel.render(output);
        output.close();
    }

    //如何合并单元格的例子，合并单元格最好从模板文件中获取
    @Test
    public void renderWithSpan() throws IOException, InvocationTargetException, IllegalAccessException {
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
        userList.add(new SalaryPayment("user1", 10000d, 10d, 1000d, 2d, 3d, 4d, 5d));
        userList.add(new SalaryPayment("user2", 10000d, 10d, 1000d, 2d, 3d, 4d, 5d));
        userList.add(new SalaryPayment("user3", 10000d, 10d, 1000d, 2d, 3d, 4d, 5d));
        userList.add(new SalaryPayment("user4", 10000d, 10d, 1000d, 2d, 3d, 4d, 5d));


        sheet.setData(Arrays.asList("userName", "baseSalary", "fullAttendanceBonus", "mealSupplement", "transportationAllowance", "sickLeave", "personalLeave", "actualPay"),
                userList);


        excel.addSheet(sheet);

        FileOutputStream output = new FileOutputStream("e:/test2.xls");
        excel.render(output);
        output.close();
    }

    //从模板文件引入表头
}
