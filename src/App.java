import java.io.*;
import java.util.*;
import java.util.List;

import javax.swing.JFileChooser;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import com.opencsv.CSVWriter;

public class App {

    public static void main(String[] args) throws FileNotFoundException, IOException {
        File csvfile = new File("src\\csvfile.csv");

        JFileChooser window = new JFileChooser();
        int returnValue = window.showOpenDialog(null);

        if(returnValue == JFileChooser.APPROVE_OPTION){
            XWPFDocument docx = new XWPFDocument(new FileInputStream(window.getSelectedFile()));
            
            try {
                FileWriter outputfile = new FileWriter(csvfile);
                CSVWriter writer = new CSVWriter(outputfile);
                List<String[]> data = new ArrayList<String[]>();

                
                List<XWPFParagraph> para = docx.getParagraphs();

                for(XWPFParagraph paragraph: para ){
                    System.out.println(paragraph.getText());
                    String[] sep_data = paragraph.getText().split(" ");
                    data.add(sep_data);
                }

                writer.writeAll(data);

                docx.close();
                writer.close();
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

    }
}
