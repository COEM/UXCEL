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
            File file = new File(name);
            if (file.exists()){

                DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
                DocumentBuilder builder = factory.newDocumentBuilder();
                Document doc = builder.parse(name);
                TransformerFactory tFactory = TransformerFactory.newInstance();
                Transformer tFormer = tFactory.newTransformer();
                Element element = (Element)doc.getElementsByTagName(ElementRemove).item(0);
                element.getParentNode().removeChild(element);
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
            
        }
    }
}
