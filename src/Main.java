import javax.swing.filechooser.FileSystemView;

public class Main {

    public static void main(String[] args) {
        System.out.println("Hello World!");
        // 获取桌面路径
        FileSystemView fsv = FileSystemView.getFileSystemView();
        String desktop = fsv.getHomeDirectory().getPath();
        String filePath = new String(desktop + "\\temp\\testExcelData.xlsx");
        // 读取Excel文件内容
        ExcelReader.readExcel(filePath);

        filePath = new String(desktop + "\\temp\\template.xls");
        ExcelReader.readExcel(filePath);
    }
}
