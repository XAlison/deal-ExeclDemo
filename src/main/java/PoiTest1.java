/**
 * Created by Any on 2018/7/15.
 */
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFDataValidation;
import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.util.CellRangeAddressList;
import java.io.IOException;
import java.io.FileOutputStream;

public class PoiTest1{

    public static void main(String [] args) throws IOException {
        HSSFWorkbook wb=new HSSFWorkbook();//excel文件对象
        HSSFSheet sheetlist=wb.createSheet("sheetlist");//工作表对象
        HSSFSheet sheetview=wb.createSheet("sheetview");
        FileOutputStream out=new FileOutputStream("D://success.xls");

        //得到验证对象
        HSSFDataValidation data_validation_list = PoiTest1.setDataValidationList((short)1,(short)1,(short)2,(short)2);

        HSSFDataValidation data_validation_view = PoiTest1.setDataValidationView((short)1,(short)1,(short)2,(short)2);
        //设置提示内容,标题,内容
        data_validation_view.createPromptBox("mm","www.tangxiangcun.com");
        //工作表添加验证数据
        sheetlist.addValidationData(data_validation_list);
        sheetview.addValidationData(data_validation_view);
        wb.write(out);
        out.close();
    }

    public static HSSFDataValidation setDataValidationList(short firstRow,short firstCol,short endRow, short endCol){
        //设置下拉列表的内容
        String[] textlist={"列表1","列表2","列表3","列表4","列表5"};

        //加载下拉列表内容
        DVConstraint constraint=DVConstraint.createExplicitListConstraint(textlist);
        //设置数据有效性加载在哪个单元格上。

        //四个参数分别是：起始行、终止行、起始列、终止列
        CellRangeAddressList regions=new CellRangeAddressList(firstRow,firstCol,endRow,endCol);
        //数据有效性对象
        HSSFDataValidation data_validation_list = new HSSFDataValidation(regions, constraint);

        return data_validation_list;
    }
    public static HSSFDataValidation setDataValidationView(short firstRow,short firstCol,short endRow, short endCol){
        //构造constraint对象
        DVConstraint constraint=DVConstraint.createCustomFormulaConstraint("B1");
        //四个参数分别是：起始行、终止行、起始列、终止列
        CellRangeAddressList regions=new CellRangeAddressList(firstRow,firstCol,endRow,endCol);
        //数据有效性对象
        HSSFDataValidation data_validation_view = new HSSFDataValidation(regions, constraint);

        return data_validation_view;
    }

}
