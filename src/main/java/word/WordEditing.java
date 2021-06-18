package word;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class WordEditing {

    public static void main(String[] args) throws IOException, InvalidFormatException {

        ArrayList<String> variableList = new ArrayList<>();
        //заполнение списка переменных
        fillingVariableList(variableList);

        ArrayList<String> list = new ArrayList<>();
        //заполнение списка значений, на которые нужно поменять переменные
        fillingList(list);

        try {
            FileInputStream fileInputStream = new FileInputStream("C:/Users/m8rin/Desktop/template.docx");
            XWPFDocument docxFile = new XWPFDocument(OPCPackage.open(fileInputStream));

            //поиск переменных в таблицах и замена
            tableSearch(docxFile, variableList, list);

            //поиск переменных в абзацах и замена
            paragraphsSearch(docxFile, variableList, list);

            //сохранение полученного документа
            docxFile.write(new FileOutputStream("C:/Users/m8rin/Desktop/Out.docx"));
            docxFile.write(new FileOutputStream("Out.docx"));
            System.out.println("\nНовый файл успешно сохранен!");

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
    }

    private static void paragraphsSearch(XWPFDocument docxFile, ArrayList<String> variableList, ArrayList<String> list) {
        List<XWPFParagraph> paragraphs = docxFile.getParagraphs();
        for (XWPFParagraph p : paragraphs) {
            List<XWPFRun> runs = p.getRuns();
            if (runs != null) {
                for (XWPFRun r : runs) {
                    String text = r.getText(0);
                    replaceText(variableList, list, r, text);
                }
            }
        }
    }

    private static void tableSearch(XWPFDocument docxFile, ArrayList<String> variableList, ArrayList<String> list) {
        for (XWPFTable tbl : docxFile.getTables()) {
            for (XWPFTableRow row : tbl.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph p : cell.getParagraphs()) {
                        for (int i = 0; i < p.getRuns().size(); i++) {
                            XWPFRun r = p.getRuns().get(i);
                            String text = r.getText(0);
                            //System.out.println(text);

                            //замена текста найденной перменной
                            replaceText(variableList, list, r, text);
                        }
                    }
                }
            }
        }
    }

    private static void replaceText(ArrayList<String> variableList, ArrayList<String> list, XWPFRun r, String text) {
        for (int i = 0; i < variableList.size(); i++) {
            if (text != null && text.contains(variableList.get(i))) {
                text = text.replace(variableList.get(i), list.get(i));
                System.out.println(variableList.get(i) + " = " + text);
                r.setText(text, 0);
            }
        }
    }

    private static void fillingVariableList(ArrayList<String> list) {
        list.add("{organization}");
        list.add("{address}");
        list.add("{phone number}");
        list.add("{fax}");
        list.add("{document number}");
        list.add("{date}");
        list.add("{number}");
        list.add("{name}");
        list.add("{customer}");
        list.add("{delivery address}");
        list.add("{customer phone number}");
    }

    private static void fillingList(ArrayList<String> list) {
        list.add("ООО 'Круто'");
        list.add("г. Уфа, ул. Кольцевая, 7");
        list.add("89666665521");
        list.add("77899");
        list.add("22001");
        list.add("15.06.2021");
        list.add("1234");
        list.add("Такой-то");
        list.add("Иванов Иван Иванович");
        list.add("г. Уфа, ул. Кольцевая, 72");
        list.add("+7 967 74 77 777");
    }
}


