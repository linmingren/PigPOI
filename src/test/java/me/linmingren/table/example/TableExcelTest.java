package me.linmingren.table.example;


import me.linmingren.table.*;
import me.linmingren.table.example.model.SalaryPayment;
import me.linmingren.table.example.model.User;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellUtil;
import org.junit.Test;

import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.util.*;

import static junit.framework.TestCase.assertTrue;
import static org.apache.poi.ss.util.CellUtil.FILL_FOREGROUND_COLOR;
import static org.apache.poi.ss.util.CellUtil.FONT;

/**
 * Unit test for simple App.
 */
public class TableExcelTest {

    //最简单的例子, 默认表头样式
    @Test
    public void simpleRender() throws IOException, TableExcelException {
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
    public void customHeader() throws IOException, TableExcelException {
        TableExcel excel = new TableExcel();
        TableSheet sheet = new TableSheet("sheet1");

        TableHeaderRow row = new TableHeaderRow(){
            @Override
            protected CellStyle updatedStyle() {
                CellStyle cellStyle = excel.getWorkbook().createCellStyle();
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
    }

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
    public void customCell() throws IOException, TableExcelException {
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

    //如何合并单元格的例子，合并单元格最好从模板文件中获取
    @Test
    public void headerSpan() throws IOException, TableExcelException {
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
    }

    //从模板文件引入表头
    @Test
    public void importSimpleTableFromTemplate() throws IOException, TableExcelException {
        TableExcel excel = new TableExcel();
        excel.importHeadersFromTemplate("headerSpan.xls");

        //这里的sheet的名字必须和模板文件中的一个sheet一样，否则就不会从模板引入样式
        TableSheet sheet = new TableSheet("sheet1");
        List<SalaryPayment> userList = new ArrayList<>();
        userList.add(new SalaryPayment("老王", 100000, 0, 1000, 200, 300, 400));
        userList.add(new SalaryPayment("小明", 10000, 1000, 1000, 300, 3000, 400));
        userList.add(new SalaryPayment("超人", 20000, 0, 1000, 0, 0, 400));


        sheet.setData(Arrays.asList("userName", "baseSalary", "fullAttendanceBonus", "mealSupplement", "transportationAllowance", "sickLeave", "personalLeave", "actualPay"),
                userList);

        excel.addSheet(sheet);

        FileOutputStream output = new FileOutputStream("excels/importFromTemplate.xls");
        excel.render(output);
        output.close();
    }
    //异常的例子
    @Test
    public void exceptionRender() throws IOException {
        TableExcel excel = new TableExcel();
        TableSheet sheet = new TableSheet("sheet1");

        TableRow row = TableHeaderRow.of(Arrays.asList("姓名", "地址", "分数", "考试时间"));
        sheet.addRow(row);

        List<User> userList = new ArrayList<>();
        userList.add(new User("老王", "隔壁", 59, new Date()));
        userList.add(new User("小明", "草地上", 80, new Date()));
        userList.add(new User("超人", "飞机上", 100, new Date()));
        //createdAtd 不存在
        try {
            sheet.setData(Arrays.asList("name", "address", "score", "createdAtd"), userList);
            assertTrue("must throw exception on unknown field", false);
        } catch (TableExcelException e) {
        }

    }
}
