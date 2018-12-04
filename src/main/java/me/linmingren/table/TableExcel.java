package me.linmingren.table;

import lombok.Data;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

@Data
public class TableExcel {
    private List<TableSheet> sheets;

    HSSFWorkbook workbook; //对应的POI对象

    public TableExcel() {
        this.sheets = new ArrayList<>();
    }

    public void addSheet(TableSheet sheet) {
        this.sheets.add(sheet);
    }

    public void render(OutputStream outputStream) throws IOException {
        workbook = new HSSFWorkbook();

        for (TableSheet s :sheets) {
            HSSFSheet sheet = workbook.createSheet(s.getName());
            s.render(sheet);
        }

        workbook.write(outputStream);
    }
}
