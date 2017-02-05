/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package bin;

/**
 *
 * @author cacing
 */
import java.io.*;
import javax.swing.JOptionPane;
import org.w3c.dom.*;
import javax.xml.parsers.*;
import javax.xml.transform.*; 
import javax.xml.transform.dom.DOMSource; 
import javax.xml.transform.stream.StreamResult;

public class xc_xmlRemove {
    public static void xmlRemovePassword(String name, String ElementRemove){
        try{
            BufferedReader bf = new BufferedReader(new InputStreamReader(System.in));
            //System.out.print("Enter a XML file name: ");
//            String xmlFile = ;
            File file = new File(name);
            //System.out.print("Enter an element which have to delete: ");
            //String remElement = bf.readLine();
            if (file.exists()){

                DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
                DocumentBuilder builder = factory.newDocumentBuilder();
                Document doc = builder.parse(name);
                TransformerFactory tFactory = TransformerFactory.newInstance();
                Transformer tFormer = tFactory.newTransformer();
                Element element = (Element)doc.getElementsByTagName(ElementRemove).item(0);
//				Remove the node
                element.getParentNode().removeChild(element);
//                              Normalize the DOM tree to combine all adjacent nodes
                doc.normalize();
                Source source = new DOMSource(doc);
                final OutputStream os = new FileOutputStream(name);
                try (PrintStream printStream = new PrintStream(os)) {
                    Result dest = new StreamResult(printStream);
                    tFormer.transform(source, dest);
                }
            }
            else{
                    JOptionPane.showMessageDialog(null, "File not found!");
            }
        }
        catch (Exception e){
//            JOptionPane.showMessageDialog(null, e, "Error Removing Password", 0);
//            System.exit(0);
        }
    }
    
    public static void main(String[] args) {
        //
         //xc_xmlRemove.xmlRemovePassword(xc_sheetList.getSheetName().get("sheet1")+".xml", "sheetProtection");
    }
}
