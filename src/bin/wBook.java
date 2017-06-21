/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package bin;

import java.io.File;

/**
 *
 * @author Cacing
 */
public class wBook {
    public static File workBook(){
        return  new File(System.getProperty("java.io.tmpdir") + "var/"+uxcel.getFileName()+"/xl/workbook.xml");
    }
    
    public static void main(String[] args) {
        
        System.out.println(workBook());
    }
    
}
