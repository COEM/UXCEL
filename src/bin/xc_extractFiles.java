package bin;

/**
 *
 * @author cacing
 */
import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Enumeration;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipInputStream;
import javax.swing.JOptionPane;
 
public class xc_extractFiles {
    ClassLoader loader = uxcel.class.getClassLoader();
    public static void unzip(String strZipFile) {

            try
            {
            File fSourceZip = new File(strZipFile);
            String zipPath = strZipFile.substring(0, strZipFile.length()-4);
            File temp = new File(zipPath);
            ZipFile zipFile = new ZipFile(fSourceZip);
            Enumeration e = zipFile.entries();

            while(e.hasMoreElements())
            {
                ZipEntry entry = (ZipEntry)e.nextElement();
                File destinationFilePath = new File(System.getProperty("java.io.tmpdir") + "var/" + uxcel.getFileName(),entry.getName());

                destinationFilePath.getParentFile().mkdirs();

                if(entry.isDirectory()){
                    continue;
                } else
                {
                    BufferedInputStream bis = new BufferedInputStream(zipFile.getInputStream(entry));

                    int b;
                    byte buffer[] = new byte[1024];
                    FileOutputStream fos = new FileOutputStream(destinationFilePath);
                    BufferedOutputStream bos = new BufferedOutputStream(fos,
                                                    1024);

                    while ((b = bis.read(buffer, 0, 1024)) != -1) {
                        bos.write(buffer, 0, b);
                    }

                    bos.flush();
                    bos.close();

                    bis.close();
                }
            }
        }
            catch(IOException ioe)
        {
            JOptionPane.showMessageDialog(null, ioe);
        }

    }
}
