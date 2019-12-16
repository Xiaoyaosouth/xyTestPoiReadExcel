package util;

import java.io.PrintWriter;
import java.util.ArrayList;

public class FileUtil {

    /**
     * 将ArrayList<String>写入文件
     * @param strAryList
     * @param targetFile 写入文件的路径
     */
    public static void writeTotxtFile(ArrayList<String> strAryList, String targetFile) throws Exception{
        PrintWriter printWriter = new PrintWriter(targetFile);
        for (String jsonStr : strAryList) {
            printWriter.println(jsonStr);
            printWriter.flush();
        }
        printWriter.close();
    }

}
