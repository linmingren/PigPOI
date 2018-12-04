package me.linmingren.table;

import lombok.Data;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.ss.usermodel.*;

import java.util.ArrayList;
import java.util.List;

@Data
public class TableRow {
    protected List<TableCell> cells;
    protected Workbook workbook;

    public TableRow() {
        this.cells = new ArrayList<>();
    }

    public void addCell(TableCell cell) {
        cells.add(cell);
    }

    private int findUnusedCol(int[] freeRow, int row) {
        for (int i = 0 ;i <freeRow.length; ++i) {
           if (freeRow[i] == row) {
               return i;
           }
        }

        return -1;
    }

    public void render(int[] freeRow, Row row) {

        int nextCol = findUnusedCol(freeRow,  row.getRowNum());

        for (int i = 0; i < cells.size(); ++i) {
            TableCell tableCell = cells.get(i);
            for (int j = nextCol; j < nextCol + tableCell.getColSpan();++j) {
                freeRow[j] = freeRow[j] + tableCell.getRowSpan();
            }

            Cell cell = row.createCell(nextCol);
            //如果当前行设置了样式，则把样式应用到该行所有的单元格上
            tableCell.setRowStyle(updatedStyle());
            tableCell.render(cell, nextCol);

            nextCol += tableCell.getColSpan();
        }
    }

    //返回新的样式来渲染该行的单元格
    protected CellStyle updatedStyle() {
        return null;
    }
}
