/**
 * Created by Any on 2018/7/15.
 */


import com.sun.corba.se.spi.ior.IdentifiableFactory;
import jdk.nashorn.internal.objects.NativeUint8Array;
import org.apache.poi.hpsf.DocumentSummaryInformation;
import org.apache.poi.hpsf.SummaryInformation;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.util.CellRangeAddressList;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.UUID;

/**
 * 从数据库中读取工资的字段，然后动态生成excel模板
 *
 * @author qiulinhe
 * <p>
 * 2017年2月20日 下午5:41:35
 */
public class ExcelOutputUtil {


    /**
     * @param @param filePath  Excel文件路径
     * @param @param handers   Excel列标题(数组)
     * @param @param downData  下拉框数据(数组)
     * @param @param downRows  下拉列的序号(数组,序号从0开始)
     * @return void
     * @throws
     * @Title: createExcelTemplate
     * @Description: 生成Excel导入模板
     */
    private static void createExcelTemplate(String filePath, String title, String[] heads, String[] dataType,String[] dataValue,
                                            List<String[]> downData, String[] downRows, String[] dateRows) throws Exception {
        //创建工作薄
        HSSFWorkbook wb = new HSSFWorkbook();
        // 创建文档信息
        wb.createInformationProperties();
        // 文档信息
        DocumentSummaryInformation documentSummaryInformation = wb.getDocumentSummaryInformation();//摘要信息
        documentSummaryInformation.setCategory("类别:Excel数据模板文件");
        documentSummaryInformation.setManager("管理者:创建者");
        documentSummaryInformation.setCompany("公司:----");
        SummaryInformation summaryInformation = wb.getSummaryInformation();
        summaryInformation.setSubject("主题:--");
        summaryInformation.setTitle("标题:测试文档");
        summaryInformation.setAuthor("作者:xxx");
        summaryInformation.setComments("备注:文档");

        //表头样式
        HSSFCellStyle style = wb.createCellStyle();
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 创建一个居中格式
        style.setFillBackgroundColor(HSSFColor.SKY_BLUE.index);
        //字体样式
        HSSFFont fontStyle = wb.createFont();
        fontStyle.setFontName("微软雅黑");
        fontStyle.setFontHeightInPoints((short) 12);
        fontStyle.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
        style.setFont(fontStyle);
        //设置背景颜色
        style.setFillForegroundColor(HSSFColor.SKY_BLUE.index);
        style.setFillPattern(HSSFColor.SKY_BLUE.index);
        HSSFSheet sheet1 = wb.createSheet(title);
        // sheet1内容
        HSSFRow rowFirst = sheet1.createRow(0);
        rowFirst.setHeightInPoints(20);
        rowFirst.setHeight((short) (25 * 20));
        // 设置标题
        for (int i = 0; i < heads.length; i++) {
            HSSFCell cell1 = rowFirst.createCell(i);
            sheet1.setColumnWidth(i, 5000);
            cell1.setCellStyle(style);
            cell1.setCellValue(heads[i]);
        }
        // 设置数据类型
        HSSFRow rowTwo = sheet1.createRow(1);
        rowTwo.setHeightInPoints(20);
        rowTwo.setHeight((short) (25 * 20));
        for (int j = 0; j < dataType.length; j++) {
            HSSFCell cell2 = rowTwo.createCell(j);
            cell2.setCellStyle(style);
            cell2.setCellValue(dataType[j]);
            sheet1.setColumnWidth(j, 4000);

        }
        // 设置示例数据
        HSSFRow rowThree = sheet1.createRow(2);
        rowThree.setHeightInPoints(20);
        rowThree.setHeight((short) (25 * 20));
        for (int k = 0; k < dataValue.length; k++) {
            HSSFCell cell3 = rowThree.createCell(k);
            cell3.setCellValue(dataValue[k]);
            sheet1.setColumnWidth(k, 4000);
            cell3.setCellStyle(style);
        }
        // 设置下拉框数据以及验证
        if (downRows != null && downRows.length > 0) {

            for (int r = 0; r < downRows.length; r++) {
                String[] dlData = downData.get(r);
                int rowNum = Integer.parseInt(downRows[r]);
                HSSFDataValidation hssfDataValidation = addDataValidationBoxs((short) 2, (short) 50000, (short) rowNum, (short) rowNum, dlData,  wb,    sheet1 );
               //HSSFWorkbook workbook, HSSFSheet tarSheet, String[] menuItems, int firstRow, int lastRow, int column
                addDropDownList(wb,sheet1,dlData,2,5000,rowNum);
                sheet1.addValidationData(hssfDataValidation); //超过255个报错
            }
        }
        // 设置日期格式以及验证
        if (dateRows != null && dateRows.length > 0) {
            for (int r = 0; r < dateRows.length; r++) {
                int rowNum = Integer.parseInt(dateRows[r]);
                HSSFDataValidation hssfDataValidation = addDateValidation((short) 2, (short) 50000, (short) rowNum, (short) rowNum, "1700-01-01",
                        "5000-01-01", "yyyy-mm-dd");
                sheet1.addValidationData(hssfDataValidation);
            }
        }
        // 扩展数据类型
        try {
            File f = new File(filePath);
            if (!f.getParentFile().exists()) {
                f.getParentFile().mkdirs();
            }
            if (!f.exists()) {
                f.createNewFile();
            }
            FileOutputStream out = new FileOutputStream(f);
            out.flush();
            wb.write(out);
            out.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }






    public static void main(String[] args) {

        //模板名称
        String fileName = "D://员工信息表.xls";
        //列标题
        String[] heads = {"姓名", "性别", "证件类型", "证件号码", "服务结束时间", "参保地", "民族"};
        //列标题
        String[] dataType = {"姓名", "性别", "证件类型", "证件号码", "服务结束时间", "参保地", "民族"};
        //列标题
        String[] dataValue = {"姓名", "性别", "证件类型", "证件号码", "服务结束时间", "参保地", "民族"};
        // 下拉类型
        List<String[]> downData = new ArrayList();
        String[] str1 = {"男", "女", "未知"};
        String[] str2 = {"北京", "上海", "广州", "深圳", "武汉", "长沙", "湘潭"};
        String[] str3 = {"01-汉族", "02-蒙古族"};

        String [] str=new String[100];
        for (int i=0;i<str.length;i++)
        {
            str[i]="下拉框".concat(i+"");
        }


        downData.add(str);
        downData.add(str1);
        downData.add(str2);


        // 下拉的列序号数组(序号从0开始)
        String[] downRows = {"1", "5", "6"};
        // 日期类型
        String[] dateRows = {}; //下拉的列序号数组(序号从0开始)
        try {
            createExcelTemplate(fileName, "员工信息表", heads, dataType,dataValue, downData, downRows, dateRows);
        } catch (Exception e) {

        }

    }



    /**
     * 单元格添加下拉菜单(不限制菜单可选项个数)<br/>
     * [注意：此方法会添加隐藏的sheet，可调用getDataSheetInDropMenuBook方法获取用户输入数据的未隐藏的sheet]<br/>
     * [待添加下拉菜单的单元格 -> 以下简称：目标单元格]
     * @param @param workbook
     * @param @param tarSheet 目标单元格所在的sheet
     * @param @param menuItems 下拉菜单可选项数组
     * @param @param firstRow 第一个目标单元格所在的行号(从0开始)
     * @param @param lastRow 最后一个目标单元格所在的行(从0开始)
     * @param @param column 待添加下拉菜单的单元格所在的列(从0开始)
     */
    public static void addDropDownList(HSSFWorkbook workbook, HSSFSheet tarSheet, String[] menuItems, int firstRow, int lastRow, int column) throws Exception
    {
        if(null == workbook){
            throw new Exception("workbook为null");
        }
        if(null == tarSheet){
            throw new Exception("待添加菜单的sheet为null");
        }

        //必须以字母开头，最长为31位
        String hiddenSheetName = "a" + UUID.randomUUID().toString().replace("-", "").substring(1, 31);
        //excel中的"名称"，用于标记隐藏sheet中的用作菜单下拉项的所有单元格
        String formulaId = "form" + UUID.randomUUID().toString().replace("-", "");
        HSSFSheet hiddenSheet = workbook.createSheet(hiddenSheetName);//用于存储 下拉菜单数据
        //存储下拉菜单项的sheet页不显示
        workbook.setSheetHidden(workbook.getSheetIndex(hiddenSheet), true);

        HSSFRow row = null;
        HSSFCell cell = null;
        //隐藏sheet中添加菜单数据
        for (int i = 0; i < menuItems.length; i++)
        {
            row = hiddenSheet.createRow(i);
            //隐藏表的数据列必须和添加下拉菜单的列序号相同，否则不能显示下拉菜单
            cell = row.createCell(column);
            cell.setCellValue(menuItems[i]);
        }
        HSSFName namedCell = workbook.createName();//创建"名称"标签，用于链接
        namedCell.setNameName(formulaId);
        namedCell.setRefersToFormula(hiddenSheetName + "!A$1:A$" + menuItems.length);
        HSSFDataValidationHelper dvHelper = new HSSFDataValidationHelper(tarSheet);
        DataValidationConstraint dvConstraint = dvHelper.createFormulaListConstraint(formulaId);

        CellRangeAddressList addressList = new CellRangeAddressList(firstRow, lastRow, column, column);
        HSSFDataValidation validation = (HSSFDataValidation)dvHelper.createValidation(dvConstraint, addressList);//添加菜单(将单元格与"名称"建立关联)
        tarSheet.addValidationData(validation);
    }


  /*  *//**
     * 从调用addDropDownList后添加下拉菜单的Workbook中获取用户输入数据的shee列表
     * @param book
     * @return
     *//*
    public static List<HSSFSheet> getDataSheetInDropMenuBook(HSSFWorkbook book){
        return getUnHideSheets(book);
    }

    *//**
     * 获取所有未隐藏的sheet
     * @param book
     * @return
     *//*
    public static List<HSSFSheet> getUnHideSheets(HSSFWorkbook book){
        List<HSSFSheet> ret = new ArrayList<HSSFSheet>();
        if(null == book){
            return ret;
        }
        int sheetCnt = book.getNumberOfSheets();
        for (int i = 0; i < sheetCnt; i++) {
            if(!book.isSheetHidden(i)){
                ret.add(book.getSheetAt(i));
            }
        }
        return ret;
    }*/


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
     * @param firstRow 首行
     * @param lastRow  行末
     * @param firstCol 行单元
     * @param lastCol  列单元
     * @param array
     * @return
     */
    public static HSSFDataValidation addDataValidationBoxs(short firstRow, short lastRow, short firstCol, short lastCol, String[] array, HSSFWorkbook wb,HSSFSheet sheet1 ) {
        DVConstraint constraint=null;
        HSSFDataValidation dataValidation=null;
        if (array.length>50)
        {

            //必须以字母开头，最长为31位
            String hiddenSheetName = "a" + UUID.randomUUID().toString().replace("-", "").substring(1, 31);
            //excel中的"名称"，用于标记隐藏sheet中的用作菜单下拉项的所有单元格
            String formulaId = "form" + UUID.randomUUID().toString().replace("-", "");
            HSSFSheet hiddenSheet = wb.createSheet(hiddenSheetName);//用于存储 下拉菜单数据
            //存储下拉菜单项的sheet页不显示
            wb.setSheetHidden(wb.getSheetIndex(hiddenSheet), false);

            HSSFRow row ;
            HSSFCell cell;
            //隐藏sheet中添加菜单数据
            for (int i = 0; i < array.length; i++)
            {
                row = hiddenSheet.createRow(i);
                //隐藏表的数据列必须和添加下拉菜单的列序号相同，否则不能显示下拉菜单
                cell = row.createCell(i);
                cell.setCellValue(array[i]);
            }
            HSSFName namedCell = wb.createName();//创建"名称"标签，用于链接
            namedCell.setNameName(formulaId);
            namedCell.setRefersToFormula(hiddenSheetName + "!B$1:B$" + array.length);
            HSSFDataValidationHelper dvHelper = new HSSFDataValidationHelper(sheet1);
            DataValidationConstraint dvConstraint = dvHelper.createFormulaListConstraint(formulaId);
            CellRangeAddressList regions = new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol);
            dataValidation = (HSSFDataValidation)dvHelper.createValidation(dvConstraint, regions);//添加菜单(将单元格与"名称"建立关联)
            wb.setSheetHidden(1, true); // 1隐藏、0显示

        }

        else
        {
            constraint = DataValidationUtil
                    .getListDVConstraint(array);
            CellRangeAddressList regions = new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol);
            dataValidation = new HSSFDataValidation(regions, constraint);


        }
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
     * @param firstRow   首行
     * @param lastRow    行末
     * @param firstCol   行单元
     * @param lastCol    列单元
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
