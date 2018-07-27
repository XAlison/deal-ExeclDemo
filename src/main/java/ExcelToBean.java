

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;


/**
 * @version V1.0
 * @Description:解析execl
 * @Author xiewl
 * @Date in 18:27 on 2018/03/26.
 */
public class ExcelToBean {
    private Workbook wb;
    private Sheet sheet;
    private Row row;

    public ExcelToBean(String filePath) {
        if (filePath == null) {
            return;
        }
        String ext = filePath.substring(filePath.lastIndexOf("."));
        try {
            InputStream is = new FileInputStream(filePath);
            if (".xls".equals(ext)) {
                wb = new HSSFWorkbook(is);
            } else if (".xlsx".equals(ext)) {
                wb = new XSSFWorkbook(is);
            } else {
                wb = null;
            }
        } catch (FileNotFoundException e) {

        } catch (IOException e) {

        }
    }

    /**
     * param:
     * Describe: 获取标题
     * Author: xiewl
     *
     * @Date in 19:04 on 2018/03/26.
     * @version V1.0
     **/
    public String[] readExcelTitle() throws Exception {
        if (wb == null) {
            throw new Exception("Workbook对象为空！");
        }
        sheet = wb.getSheetAt(0);
        row = sheet.getRow(0);
        // 标题列数
        int colNum = row.getPhysicalNumberOfCells();
        String[] title = new String[colNum];
        for (int i = 0; i < colNum; i++) {
            title[i] = row.getCell(i).getStringCellValue();
        }
        return title;
    }


    /**
     * Describe: 获取内容
     * Author: xiewl
     *
     * @Date in 19:05 on 2018/03/26.
     * @version V1.0
     **/
    public Map<Integer, Map<String, Object>> readExcelContent() throws Exception {
        if (wb == null) {
            throw new Exception("Workbook对象为空！");
        }
        Map<Integer, Map<String, Object>> content = new HashMap<Integer, Map<String, Object>>();
        sheet = wb.getSheetAt(0);
        // 总行数
        int rowNum = sheet.getLastRowNum();
        row = sheet.getRow(0);
        int colNum = row.getPhysicalNumberOfCells();
        String[] columns = readExcelTitle();
        // 第一行为表头的标题
        for (int i = 1; i <= rowNum; i++) {
            row = sheet.getRow(i);
            if (row == null) {
                sheet.shiftRows(i + 1, sheet.getLastRowNum(), -1);
            }
            int j = 0;
            Map<String, Object> cellValue = new HashMap<String, Object>();
            while (j < colNum) {
                Object obj = getCellFormatValue(row.getCell(j));
                cellValue.put(columns[j], obj);
                j++;
            }
            content.put(i, cellValue);
        }
        return content;
    }

    /**
     * param: Cell
     * Describe: 单元格值
     * Author: xiewl
     *
     * @Date in 19:05 on 2018/03/26.
     * @version V1.0
     **/
    private Object getCellFormatValue(Cell cell) {
        String value = "";
        if (cell != null) {
            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) { //日期类型
                        Date date = cell.getDateCellValue();
                        value = "";
                    } else {
                        Double data = cell.getNumericCellValue();
                        value = data.toString();
                    }
                    break;
                case Cell.CELL_TYPE_STRING:
                    value = cell.getStringCellValue();
                    break;
                case Cell.CELL_TYPE_BOOLEAN:
                    Boolean data = cell.getBooleanCellValue();
                    value = data.toString();
                    break;
                case Cell.CELL_TYPE_ERROR:
                    value = "";
                    break;
                case Cell.CELL_TYPE_FORMULA:
                    value = String.valueOf(cell.getNumericCellValue());
                    if (value.equals("NaN")) {
                        value = cell.getStringCellValue().toString();
                    }
                    break;
                case Cell.CELL_TYPE_BLANK:
                    value = "";
                    break;
                default:
                    value = cell.getStringCellValue().toString();
                    break;
            }
        }
        return value;
    }
}
