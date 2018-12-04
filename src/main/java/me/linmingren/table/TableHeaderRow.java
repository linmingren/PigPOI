package me.linmingren.table;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.util.List;

public class TableHeaderRow extends TableRow {
    @Override
    protected CellStyle updatedStyle( ) {
        return defaultHeaderRowStyle();
    }

    public static TableHeaderRow of(List<String> title) {
        TableHeaderRow row = new TableHeaderRow();
        for (String t :title) {
            row.addCell(new TableCell(t));
        }

        return row;
    }

    private CellStyle defaultHeaderRowStyle() {
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setAlignment(HorizontalAlignment.CENTER.CENTER);


        Font f = workbook.createFont();
        f.setBold(true);
        cellStyle.setFont(f);

        return cellStyle;
    }
}
