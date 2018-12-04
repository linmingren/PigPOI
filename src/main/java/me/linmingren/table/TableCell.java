package me.linmingren.table;

import lombok.Data;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;

import java.math.BigDecimal;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

@Data
public class TableCell {
    private  Object value;

    private int rowSpan = 1;
    private int colSpan = 1;
    //所属于行的样式
    private CellStyle rowStyle;
    //只支持数字，字符串和布尔类型，其他类型会直接当做字符串处理
    public TableCell(Object value) {
        this.value = value;
    }

    public void render(HSSFCell cell, int nextCol) {
        setCellValue(cell);

        //如果设置了行样式，则继承
        if (rowStyle != null) {
            cell.setCellStyle(rowStyle);
        }

        //如果子类更新了样式，则覆盖对应的样式
        CellUtil.setCellStyleProperties(cell, updatedStyle());

        //合并单元格
        HSSFRow row = cell.getRow();

        if (rowSpan >1) {
            CellRangeAddress mergedRegion = new CellRangeAddress(row.getRowNum(),row.getRowNum() + rowSpan - 1,nextCol,nextCol);
            row.getSheet().addMergedRegion(mergedRegion);
        }

        if (colSpan >1) {
            CellRangeAddress mergedRegion = new CellRangeAddress(row.getRowNum(),row.getRowNum() ,nextCol,nextCol  + colSpan - 1);
            row.getSheet().addMergedRegion(mergedRegion);
        }
    }

    private void setCellValue(HSSFCell cell) {
        if (value instanceof BigDecimal) {
            double intValue = ((BigDecimal) value).doubleValue();
            cell.setCellValue(intValue);
        } else if (value instanceof Integer) {
            int intValue = (Integer) value;
            cell.setCellValue(intValue);
        } else if (value instanceof Float) {
            float fValue = (Float) value;
            cell.setCellValue(fValue);
        } else if (value instanceof Double) {
            double dValue = (Double) value;
            cell.setCellValue(dValue);
        } else if (value instanceof Long) {
            long longValue = (Long) value;
            cell.setCellValue(longValue);
        } else if (value instanceof Boolean) {
            boolean bValue = (Boolean) value;
            cell.setCellValue(bValue);
        } else if (value instanceof Date) {
            Date date = (Date) value;
            cell.setCellValue(date);
        } else {
            cell.setCellValue(value.toString());
        }
    }

    protected Map updatedStyle() {
       return new HashMap();
    }
}
