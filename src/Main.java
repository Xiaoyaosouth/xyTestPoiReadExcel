import excel.ExcelReader;
import util.DownloadFile;

import javax.swing.filechooser.FileSystemView;
import java.net.URL;

public class Main {

    public static void main(String[] args) {
        System.out.println("Hello World!");
        try {
            // 获取桌面路径
            FileSystemView fsv = FileSystemView.getFileSystemView();
            String desktop = fsv.getHomeDirectory().getPath();

            /*
            String filePath1 = desktop + "\\temp\\testExcelData.xlsx";
            // 读取Excel文件内容
            excel.ExcelReader.readExcel(filePath1);

            String filePath2 =desktop + "\\temp\\template.xls";
            excel.ExcelReader.readExcel(filePath2);

            // 转成JSON写入文件
            ArrayList<String> strAryList1 = excel.ExcelReader.convertExcelToJson(filePath1);
            util.FileUtil.writeTotxtFile(strAryList1, desktop + "\\temp\\testExcelData.json");
            ArrayList<String> strAryList2 = excel.ExcelReader.convertExcelToJson(filePath2);
            util.FileUtil.writeTotxtFile(strAryList2, desktop + "\\temp\\template.json");
            */
            URL url = new URL("http://xybucket.obs.cn-south-1.myhuaweicloud.com/uistudio/upload/userData.xlsx");
            DownloadFile.downloadFile(url, desktop + "\\temp\\myUserData.xlsx");
            String filePath3 = desktop + "\\temp\\myUserData.xlsx";
            ExcelReader.readExcel(filePath3);

        }catch (Exception e){
            e.printStackTrace();
        }
    }
}
