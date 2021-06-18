package excel;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

public class ExcelEditing {

    public static void main(String[] args) throws IOException {
        ArrayList<String> variableList = new ArrayList<>();
        //заполнение списка переменных
        fillingVariableList(variableList);

        ArrayList<String> list = new ArrayList<>();
        //заполнение списка значений, на которые нужно поменять переменные
        fillingList(list);

        FileInputStream file = new FileInputStream("C:/Users/m8rin/Desktop/template.xls");
        HSSFWorkbook workbook = new HSSFWorkbook(file);

        // говорим, что хотим работать с первым листом
        HSSFSheet sheet = workbook.getSheetAt(0);

        //проходим по всему листу
        for (Row row : sheet) {
            for (Cell cell : row) {
                int cellType = cell.getCellType();
                if (cellType == Cell.CELL_TYPE_STRING) {
                    for (int i = 0; i < variableList.size(); i++) {
                        if (cell.getStringCellValue().equals(variableList.get(i))) {
                            System.out.println(cell + " = " + list.get(i));
                            cell.setCellValue(list.get(i));
                        }
                    }
                }
            }
        }

        // сохранение полученного документа
        try (FileOutputStream out = new FileOutputStream("C:/Users/m8rin/Desktop/Out.xls")) {
            workbook.write(out);
        } catch (IOException e) {
            e.printStackTrace();
        }

        System.out.println("Excel файл успешно создан!");
    }

    private static void fillingVariableList(ArrayList<String> list) {
        list.add("{organization}");
        list.add("{address}");
        list.add("{numb}");
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
