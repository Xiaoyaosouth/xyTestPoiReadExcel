import javax.swing.filechooser.FileSystemView;
import java.util.ArrayList;

public class Main {

    public static void main(String[] args) {
        System.out.println("Hello World!");
        try {
            // 获取桌面路径
            FileSystemView fsv = FileSystemView.getFileSystemView();
            String desktop = fsv.getHomeDirectory().getPath();

            String filePath1 = new String(desktop + "\\temp\\testExcelData.xlsx");
            // 读取Excel文件内容
            ExcelReader.readExcel(filePath1);

            String filePath2 = new String(desktop + "\\temp\\template.xls");
            ExcelReader.readExcel(filePath2);

            // 转成JSON写入文件
            ArrayList<String> strAryList1 = ExcelReader.convertExcelToJson(filePath1);
            FileUtil.writeTotxtFile(strAryList1, desktop + "\\temp\\testExcelData.json");
            ArrayList<String> strAryList2 = ExcelReader.convertExcelToJson(filePath2);
            FileUtil.writeTotxtFile(strAryList1, desktop + "\\temp\\template.json");
        }catch (Exception e){
            e.printStackTrace();
        }
    }
}
