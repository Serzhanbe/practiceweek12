import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;
import java.util.ArrayList;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
//serzhanbejk
public class Main {
    public static void main(String[] args) {
        String filePath = "C:\\Users\\38214\\IdeaProjects\\Aldik\\src\\studentsss.xlsx";

        List<Student> students = readStudentsFromExcel(filePath);

        for (Student student : students) {
            System.out.printf("Имя: %s, Текущая стипендия: %.2f, Новая стипендия: %.2f, Увеличение: %.2f%n",
                    student.getName(),
                    student.getCurrentScholarship(),
                    student.getNewScholarship(),
                    student.getScholarshipIncrease());
        }
    }

    public static List<Student> readStudentsFromExcel(String filePath) {
        List<Student> students = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(new File(filePath));
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);

                String name = row.getCell(0).getStringCellValue();
                double currentScholarship = row.getCell(1).getNumericCellValue();
                double newScholarship = row.getCell(2).getNumericCellValue();

                students.add(new Student(name, currentScholarship, newScholarship));
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

        return students;
}
}//Эта программа на Java читает данные о студентах из файла Excel используя библиотеку Apache POI и разные библиотеки Затем она выводит информацию о каждом студенте: его имя, текущую стипендию, новую стипендию и разницу между ними
//Метод readStudentsFromExcel открывает файл с помощью FileInputStream и читает его с использованием библиотеки Apache POI
//Программа берет данные с первого листа Excel-файла
//на 11 строке мы указываем путь к экзел файлу
//В главном методе программы (main) она вызывает readStudentsFromExcel чтобы получить список студентов а затем выводит данные о каждом студенте в читаемом формате
//
