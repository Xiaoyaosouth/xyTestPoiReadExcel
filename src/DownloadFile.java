import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.net.URLConnection;

/**
 * 下载文件
 *
 * @author rk
 * 2019.12.16 10:55
 */
public class DownloadFile {
    /**
     * 下载网络文件
     *
     * @param url      http地址（BUG：使用https会报错）
     * @param filePath 保存文件路径
     */
    public static void downloadFile(URL url, String filePath) {
        int bytesum = 0;
        int byteread = 0;
        InputStream inStream = null;
        FileOutputStream fs = null;
        try {
            URLConnection conn = url.openConnection();
            inStream = conn.getInputStream();
            fs = new FileOutputStream(filePath);
            byte[] buffer = new byte[1204];
            while ((byteread = inStream.read(buffer)) != -1) {
                bytesum += byteread;
                //System.out.println(bytesum);
                fs.write(buffer, 0, byteread);
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                fs.close();
                inStream.close();
            } catch (IOException ioe) {
                ioe.printStackTrace();
            }
        }
    }

}
