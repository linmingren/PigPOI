package me.linmingren.table;

import lombok.Data;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

@Data
public class TableExcel {
    private List<TableSheet> sheets;

    Workbook workbook; //对应的POI对象

    public TableExcel() {
        this.sheets = new ArrayList<>();
        workbook = new HSSFWorkbook();
    }

    public void addSheet(TableSheet sheet) {
        this.sheets.add(sheet);
    }

    public void render(OutputStream outputStream) throws IOException {
        for (TableSheet s :sheets) {
            Sheet sheet = workbook.getSheet(s.getName());
            if (sheet == null) {
                sheet = workbook.createSheet(s.getName());
            }

            s.render(sheet);
        }

        workbook.write(outputStream);
    }

    //从模板文件导入表头

    /**
     *
     * @param templateFileName 文件名，必须存放在classpath://templates目录下
     */
    public void importHeadersFromTemplate(String templateFileName) throws IOException {
         workbook = WorkbookFactory.create(this.getClass().getResourceAsStream("/templates/" + templateFileName));
    }
}
