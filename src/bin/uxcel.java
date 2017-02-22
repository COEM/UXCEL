package bin;

import form.home;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URL;


/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

/**
 *
 * @author cacing
 */
public class uxcel {
    
    private static String fileNameExt,fileLocation,fileName,saveFileLocation;
    public static String getFileLocation(){
       return fileLocation;
    }

    public static String setFileLocation(String file){
        return fileLocation = file;
    }

    public static String getFileNameExt(){
        return fileNameExt;
    }

    public static String setFileNameExt(String name){
       return fileNameExt = name;
    }
    
    public static String setFileName(String name){
       return fileName = name;
    }
    
    public static String getFileName(){
        int pos = fileNameExt.lastIndexOf(".");
        if (pos > 0) {
            fileName = fileNameExt.substring(0, pos);
        }
        return fileName;
    }
    
    public static void copyExcel(File source, File dest) throws IOException {
        InputStream inStream = null;
    	OutputStream outStream = null;
    	try{
// 
//    	    File file1 =new File(uxcel.getFileLocation());
//    	    File file2 =new File("tmp/"+uxcel.getFileNameExt());
 
    	    inStream = new FileInputStream(source);
    	    outStream = new FileOutputStream(dest); // for override file content
    	    //outStream = new FileOutputStream(file2,<strong>true</strong>); // for append file content
 
    	    byte[] buffer = new byte[1024];
 
    	    int length;
    	    while ((length = inStream.read(buffer)) > 0){
    	    	outStream.write(buffer, 0, length);
    	    }
 
    	    if (inStream != null)inStream.close();
    	    if (outStream != null)outStream.close();
 
    	    System.out.println("File Copied..");
    	}catch(IOException e){
    		e.printStackTrace();
    	}
    }
    public static String getCleanPath() {
        URL location = uxcel.class.getProtectionDomain().getCodeSource()
                .getLocation();
        String path = location.getFile();

        return new File(path).getParent();
    }
    
    public static String setSaveLocation(String x){
        return saveFileLocation = x;
    }
    
    public static String getSaveLocation(){
        return saveFileLocation;
    }
    
    public static boolean deleteDir(File dir) {
      if (dir.isDirectory()) {
         String[] children = dir.list();
         for (int i = 0; i < children.length; i++) {
            boolean success = deleteDir (new File(dir, children[i]));
            
            if (!success) {
               return false;
            }
         }
      }
      return dir.delete();
      //System.out.println("The directory is deleted.");
   }
    
    
}
