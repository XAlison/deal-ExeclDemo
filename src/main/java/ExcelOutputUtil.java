/**
 * Created by Any on 2018/7/15.
 */


import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hpsf.DocumentSummaryInformation;
import org.apache.poi.hpsf.SummaryInformation;
import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.hssf.usermodel.HSSFDataValidation;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellRangeAddressList;

/**
 * 从数据库中读取工资的字段，然后动态生成excel模板
 *
 * @author qiulinhe
 *         <p>
 *         2017年2月20日 下午5:41:35
 */
public class ExcelOutputUtil {
    private static HSSFSheet sheet;

    public static void main(String[] args) {
        FileOutputStream out = null;
        try {
            // excel对象
            HSSFWorkbook wb = new HSSFWorkbook();
            wb.createInformationProperties();//创建文档信息
            DocumentSummaryInformation dsi = wb.getDocumentSummaryInformation();//摘要信息
            dsi.setCategory("类别:Excel数据模板文件");//类别
            dsi.setManager("管理者:创建者");//管理者
            dsi.setCompany("公司:----");//公司
            SummaryInformation si = wb.getSummaryInformation();//摘要信息
            si.setSubject("主题:--");//主题
            si.setTitle("标题:测试文档");//标题
            si.setAuthor("作者:谢文林");//作者
            si.setComments("备注:POI测试文档");//备注
            // sheet对象
            sheet = wb.createSheet("AAAA");
            // 输出excel对象
            out = new FileOutputStream("D://ceshi.xls");


            // 取得规则
            HSSFDataValidation validateData = ExcelOutputUtil.addDataValidationBoxs((short) 1, (short) 65535, (short) 1, (short) 1, new String[]{"A", "C", "D", "F"});
            HSSFDataValidation validateData1 = ExcelOutputUtil.addDateValidation((short) 1, (short) 65535, (short) 1, (short) 2, "1900-01-01",
                    "5000-01-01", "yyyy-mm-dd");
            HSSFDataValidation validateData2 = ExcelOutputUtil.addDataValidationInt((short) 1, (short) 65535, (short) 1, (short) 3, "10", "50");
            // 设定规则
            sheet.addValidationData(validateData);
            sheet.addValidationData(validateData1);
            sheet.addValidationData(validateData2);
            wb.write(out);
            out.close();

            System.out.println("在D盘成功生成了excel，请去查看");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (out != null) {
                try {
                    out.close();
                } catch (IOException e) {
                    // TODO Auto-generated catch block
                    e.printStackTrace();
                }
            }
        }
    }


    /**
     * 数字校验
     *
     * @param firstRow     首行
     * @param lastRow      行末
     * @param firstCol     行单元
     * @param lastCol      列单元
     * @param beginDecimal 最小值
     * @param endDecimal   最大值
     * @return
     */
    public static HSSFDataValidation addDataValidationDecimal(short firstRow, short lastRow, short firstCol, short lastCol, String beginDecimal, String endDecimal) {
        DVConstraint constraint = DataValidationUtil
                .getDecimalDVConstraintBetween(beginDecimal, endDecimal);
        CellRangeAddressList regions = new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol);
        // 数据有效性对象
        HSSFDataValidation dataValidation = new HSSFDataValidation(regions,
                constraint);
        dataValidation.createPromptBox("输入提示", "请填写指定格式");
        dataValidation.createErrorBox("数据格式错误提示", "介于" + beginDecimal + "-" + endDecimal + "之间");
        return dataValidation;
    }
    /**
     * @param firstRow     首行
     * @param lastRow      行末
     * @param firstCol     行单元
     * @param lastCol      列单元
     * @param beginIntData
     * @param endIntData
     * @return
     */
    public static HSSFDataValidation addDataValidationInt(short firstRow, short lastRow, short firstCol, short lastCol, String beginIntData,
                                                          String endIntData) {
        DVConstraint constraint = DataValidationUtil.getIntDVConstraintBetween(
                beginIntData, endIntData);
        CellRangeAddressList regions = new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol);
        // 数据有效性对象
        HSSFDataValidation dataValidation = new HSSFDataValidation(regions,
                constraint);
        dataValidation.createPromptBox("输入提示", "请填写数字格式");
        dataValidation.createErrorBox("数据格式错误提示", "介于" + beginIntData + "-" + endIntData + "之间");
        return dataValidation;
    }


    /**
     * @param firstRow     首行
     * @param lastRow      行末
     * @param firstCol     行单元
     * @param lastCol      列单元
     * @param array
     * @return
     */
    public static HSSFDataValidation addDataValidationBoxs(short firstRow, short lastRow, short firstCol, short lastCol, String[] array) {
        DVConstraint constraint = DataValidationUtil
                .getListDVConstraint(array);
        CellRangeAddressList regions = new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol);
        HSSFDataValidation dataValidation = new HSSFDataValidation(regions, constraint);
        // 设置只能下拉选择
        dataValidation.setEmptyCellAllowed(false);
        dataValidation.setShowPromptBox(true);
        dataValidation.createPromptBox("输入提示", "请从下拉列表中选择");
        // 设置输入错误提示信息
        dataValidation.createErrorBox("格式错误提示", "请从下拉列表中选择！");
        dataValidation.setSuppressDropDownArrow(false);
        return dataValidation;
    }

    /**
     * @param firstRow     首行
     * @param lastRow      行末
     * @param firstCol     行单元
     * @param lastCol      列单元
     * @param beginDate
     * @param endDate
     * @param dateFormat
     * @return
     */
    public static HSSFDataValidation addDateValidation(short firstRow, short lastRow, short firstCol, short lastCol, String beginDate,
                                                       String endDate, String dateFormat) {
        DVConstraint constraint = DataValidationUtil
                .getDateDVConstraintBetween(beginDate, endDate, dateFormat);
        CellRangeAddressList regions = new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol);
        // 数据有效性对象
        HSSFDataValidation dataValidation = new HSSFDataValidation(regions,
                constraint);
        dataValidation.setSuppressDropDownArrow(false);
        dataValidation.createPromptBox("输入提示", "请填写日期格式");
        // 设置输入错误提示信息
        dataValidation.createErrorBox("日期格式错误提示", "你输入的日期格式不符合'yyyy-mm-dd'格式规范，请重新输入！");
        dataValidation.setShowPromptBox(true);
        return dataValidation;
    }


}
