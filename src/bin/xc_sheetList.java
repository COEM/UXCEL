package bin;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Hashtable;
import java.util.Map;
import java.util.TreeMap;
import java.util.stream.Collectors;

/**
 *
 * @author cacing
 */
public class xc_sheetList {
    private static Map<String, String> nama = new HashMap<String, String>();
    public static Map<String, String> getSheetName(){
        File dir = new File(System.getProperty("java.io.tmpdir") + "var/"+uxcel.getFileName()+"/xl/worksheets");
        File[] listDir = dir.listFiles();
        
        for (File file : listDir) {
            if (file.isFile()) {
                String name = file.getName();
                int pos = name.lastIndexOf(".");
                int i = 1;
                if (pos > 0) {
                    name = name.substring(0, pos);
                    nama.put(name, dir + "\\"+name);
                }
            }
        }
        return nama;
    }
}
