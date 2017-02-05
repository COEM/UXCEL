/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
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
    //private static ArrayList<String, String> nama = new ArrayList<String, String>();

    public static Map<String, String> getSheetName(){
        File dir = new File(uxcel.getCleanPath() +"/var/"+uxcel.getFileName()+"/xl/worksheets");
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
    
    public static void main(String[] args) {
        System.out.println(getSheetName());
    }
}
