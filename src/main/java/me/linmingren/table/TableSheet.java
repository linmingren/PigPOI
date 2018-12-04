package me.linmingren.table;

import lombok.Data;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellUtil;

import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Data
@Slf4j
public class TableSheet {
    private String name;
    private List<TableRow> rows;

    private List data;
    private List<String> dataFieldNames;
    private Sheet templateSheet;

    public TableSheet(String name) {
        this.rows = new ArrayList<>();
        this.name = name;
    }

    int findUnusedRow(int[] freeRow) {
        int row = Integer.MAX_VALUE;
        for (int i = 0 ; i< freeRow.length; ++i) {
            if (freeRow[i] < row) {
                row = freeRow[i];
            }
        }

        return row;
    }

    public void render(Sheet sheet) {
        //下标是列的编号， 值是该列上的可用的行号， 比如freeRow[0]是0, 说明第一列上第一行是可以添加新行的
        int[] freeRow = new int[dataFieldNames.size()];
        for (int i = 0 ; i < freeRow.length; ++i) {
            freeRow[i] = sheet.getLastRowNum() > 0 ? sheet.getLastRowNum() + 1 : 0;
        }

        for (int i = 0; i < rows.size(); ++i) {
            Row row = sheet.createRow(findUnusedRow(freeRow));
            rows.get(i).setWorkbook(sheet.getWorkbook());
            rows.get(i).render(freeRow,row);
        }
    }

    public void addRow(TableRow row) {
        rows.add(row);
    }

    public void setData(List<String> fieldNames, List dataList) throws TableExcelException  {
        this.dataFieldNames = fieldNames;
        for (Object data :dataList ) {
            addData(data);
        }
    }

    public void addData(Object data) throws TableExcelException {
        Method[] methods = data.getClass().getMethods();
        Map<String, Method> methodMap = new HashMap<>();
        for (Method m : methods) {
            methodMap.put(m.getName().toLowerCase(),m);
        }

        TableRow row = new TableRow();

        for (int i = 0; i < dataFieldNames.size(); ++i) {
            ///需要通过get方法来获取字段的值，因为有些字段是没有值的，必须通过get方法才会触发计算
            String fieldName = dataFieldNames.get(i);
            Method method = methodMap.get("get" + dataFieldNames.get(i).toLowerCase());
            if (method != null) {
                try {
                    row.addCell(createDataCell(dataFieldNames.get(i),method.invoke(data), this.rows.size(),i));
                } catch (IllegalAccessException e) {
                    throw new TableExcelException("failed to get value of field:[" + fieldName +"]",e);
                } catch (InvocationTargetException e) {
                    throw new TableExcelException("failed to get value of field:[" + fieldName +"]",e);
                }
            } else {
                throw new TableExcelException("field:[" + fieldName +"] not found");
            }
        }

        this.addRow(row);
    }

    protected TableCell createDataCell(String fieldName,Object value, int row, int col) {
        return new TableCell(value);
    }
}
