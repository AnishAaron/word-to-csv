import java.io.*;
import java.util.*;
import java.util.List;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import com.opencsv.CSVWriter;

public class App {

    public static void main(String[] args) {
        File csvfile = new File("src\\csvfile.csv");
        try {
            FileWriter outputfile = new FileWriter(csvfile);
            CSVWriter writer = new CSVWriter(outputfile);
            List<String[]> data = new ArrayList<String[]>();

            FileInputStream docfile = new FileInputStream("src\\test.docx");
            XWPFDocument docx = new XWPFDocument(docfile);
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
