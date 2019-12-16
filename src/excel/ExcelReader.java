package excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.DecimalFormat;
import java.util.*;
import java.util.logging.Logger;

/**
 * 用于读取Excel文件的类
 * 【注】
 * （1）读取xlsx需要commons-collections4和commons-compress
 * （2）读取xls需要commons-math3
 * （3）poi-ooxml是ooxml-schemas的精简版，poi-3.7以上需要另外导入ooxml-schemas-1.1
 */
public class ExcelReader {

    private static Logger logger = Logger.getLogger(ExcelReader.class.getName()); // 日志打印类

    private static final String XLS = "xls";
    private static final String XLSX = "xlsx";

    /**
     * 根据文件后缀名类型获取对应的工作簿对象
     *
     * @param inputStream 读取文件的输入流
     * @param fileType    文件后缀名类型（xls或xlsx）
     * @return 包含文件数据的工作簿对象
     * @throws IOException
     */
    public static Workbook getWorkbook(InputStream inputStream, String fileType) throws IOException {
        Workbook workbook = null;
        if (fileType.equalsIgnoreCase(XLS)) {
            workbook = new HSSFWorkbook(inputStream);
        } else if (fileType.equalsIgnoreCase(XLSX)) {
            workbook = new XSSFWorkbook(inputStream);
        }
        return workbook;
    }

    /**
     * 将单元格内容转换为字符串
     *
     * @param cell 单元格
     * @return
     */
    private static String convertCellValueToString(Cell cell) {
        if (cell == null) {
            return null;
        }
        String returnValue = null;
        switch (cell.getCellType()) {
            case NUMERIC:   //数字
                Double doubleValue = cell.getNumericCellValue();
                // 格式化科学计数法，取一位整数
                DecimalFormat df = new DecimalFormat("0");
                returnValue = df.format(doubleValue);
                break;
            case STRING:    //字符串
                returnValue = cell.getStringCellValue();
                break;
            case BOOLEAN:   //布尔
                Boolean booleanValue = cell.getBooleanCellValue();
                returnValue = booleanValue.toString();
                break;
            case BLANK:     // 空值
                break;
            case FORMULA:   // 公式
                returnValue = cell.getCellFormula();
                break;
            case ERROR:     // 故障
                break;
            default:
                break;
        }
        return returnValue;
    }

    /**
     * 遍历表格单元格
     *
     * @param wb 工作簿
     */
    public static void readTable(Workbook wb) {
        Sheet sheet = wb.getSheetAt(0);
        for (Iterator ite = sheet.rowIterator(); ite.hasNext(); ) {
            // 每一行
            Row row = (Row) ite.next();
            System.out.println();
            for (Iterator itet = row.cellIterator(); itet.hasNext(); ) {
                // 每一个单元格
                Cell cell = (Cell) itet.next();
                System.out.print(convertCellValueToString(cell) + " ");
            }
        }
        System.out.println();
    }

    /**
     * 读取Excel文件内容
     *
     * @param fileName 要读取的Excel文件所在路径
     */
    public static void readExcel(String fileName) {
        Workbook workbook = null;
        FileInputStream inputStream = null;
        try {
            // 获取Excel后缀名
            String fileType = fileName.substring(fileName.lastIndexOf(".") + 1, fileName.length());
            // 获取Excel文件
            File excelFile = new File(fileName);
            if (!excelFile.exists()) {
                logger.warning("指定的Excel文件不存在！");
            }
            // 获取Excel工作簿
            inputStream = new FileInputStream(excelFile);
            workbook = getWorkbook(inputStream, fileType);
            readTable(workbook);
        } catch (Exception e) {
            logger.warning("解析Excel失败，文件名：" + fileName + " 错误信息：" + e.getMessage());
            //e.printStackTrace();
        } finally {
            try {
                if (null != workbook) {
                    workbook.close();
                }
                if (null != inputStream) {
                    inputStream.close();
                }
            } catch (Exception e) {
                logger.warning("关闭数据流出错！错误信息：" + e.getMessage());
            }
        }
    }

    /**
     * 将excel文件每一行数据转为的json字符串后放入ArrayList<>
     * @param excelPath Excel文件的路径（支持xls、xlsx）
     * @return ArrayList\<String\>
     * @throws Exception
     */
    public static ArrayList<String> convertExcelToJson(String excelPath) throws Exception {
        System.out.println("开始处理文件" + excelPath);
        ArrayList<String> jsonString = new ArrayList<>();
        //Excel工作簿
        Workbook wb = WorkbookFactory.create(new FileInputStream(excelPath));
        //sheet表格
        Sheet sheet = wb.getSheetAt(0);
        //标题
        String[] headNames = null;
        //每行
        for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
            Row row = sheet.getRow(i);
            System.out.println("——第" + (i + 1) + "行——");
            if (i == 0) { // 第1行一般是标题
                headNames = new String[row.getPhysicalNumberOfCells()];
                //每列
                for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
                    Cell cell = row.getCell(j); // 第j列
                    //headNames[j] = cell.getStringCellValue(); // 取第1行第j列的值
                    headNames[j] = convertCellValueToString(cell);
                    System.out.print(convertCellValueToString(cell) + " ");
                } // for end 获取标题结束
                System.out.println();
            }
            // 当不是第1行
            else {
                StringBuilder stringBuilder = new StringBuilder(); // 用于保存JSON字符串
                stringBuilder.append("{");
                // 开始逐行逐列构造JSON
                for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
                    Cell cell = row.getCell(j); // 当前单元格
                    String key = headNames[j]; // 当前单元格value的key {"key":value}
                    // String value = cell.getStringCellValue().replaceAll("[\\n\\r\"]", " "); // 取值并删除换行
                    String value = convertCellValueToString(cell).replaceAll("[\\n\\r\"]", " ");
                    System.out.print(value + " ");
                    // {"key":value,"key":value},
                    if (j != row.getPhysicalNumberOfCells() - 1) {
                        // 如果不是第i行的最后一列，则添加","
                        stringBuilder.append("\"").append(key).append("\"").append(":").append("\"").append(value).append("\"").append(",");
                    } else {
                        stringBuilder.append("\"").append(key).append("\"").append(":").append("\"").append(value).append("\"");
                    }
                } // for end 构造一行json数据完成
                System.out.println();
                stringBuilder.append("},"); // 每一行json尾部的"},"
                jsonString.add(stringBuilder.toString()); // 将数据放入ArrayList<String>
            }
        } // for end 读取每一行完成
        wb.close(); // 关闭工作簿
        return jsonString;
    }

}
