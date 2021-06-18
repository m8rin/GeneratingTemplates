package word;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.FileInputStream;

public class WordReader {
    public static void main(String[] args) {
        try (FileInputStream fileInputStream = new FileInputStream("C:/Users/m8rin/Desktop/Test.docx")) {

            // открываем файл и считываем его содержимое в объект XWPFDocument
            XWPFDocument docxFile = new XWPFDocument(OPCPackage.open(fileInputStream));
            XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(docxFile);
            /*
            // считываем верхний колонтитул (херед документа)
            XWPFHeader docHeader = headerFooterPolicy.getDefaultHeader();
            System.out.println(docHeader.getText());

            // печатаем содержимое всех параграфов документа в консоль
            List<XWPFParagraph> paragraphs = docxFile.getParagraphs();
            for (XWPFParagraph p : paragraphs) {
                System.out.println(p.getText());
            }
            // считываем нижний колонтитул (футер документа)
            XWPFFooter docFooter = headerFooterPolicy.getDefaultFooter();
            System.out.println(docFooter.getText());*/

            // печатаем все содержимое Word файла
            XWPFWordExtractor extractor = new XWPFWordExtractor(docxFile);
            System.out.println(extractor.getText());

        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }
}
