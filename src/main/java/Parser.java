import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.select.Elements;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class Parser {

    public static void main(String[] args) throws IOException {

        Map<String, String> map = new TreeMap<>();

        String mainUrl = "https://www.kinopoisk.ru";
        String url = "https://www.kinopoisk.ru/premiere/ru/";

        Parser.parsingWithJSON(url, map, mainUrl);
        System.out.println("Список премьер с сайта " + mainUrl + " получен.");

        Date date = new Date();
        System.out.println("Текущая дата: " + date);

        System.out.println("Для вывода нажмите:");
        System.out.println("\"1\" для вывода в КОНСОЛЬ");
        System.out.println("\"2\" для вывода в MS EXEL");

        Scanner i = new Scanner(System.in);
        int num = i.nextInt();

        switch (num) {
            case 1:
                Set<Map.Entry<String, String>> set = map.entrySet();
                for (Map.Entry<String, String> entry : set) {
                    System.out.println(entry.getKey() + ": " + entry.getValue());

                }
                break;
            case 2:
                File file = Parser.createFile("\\NewFilm.xls");
                Parser.writeIntoExcel(file, map);
                break;
            default:
                System.out.println("Некорректный ввод!");
                System.out.println("Программа завершена.");
        }
        i.close();
    }


    public static Map<String, String> parsingWithJSON(String url, Map<String, String> map, String mainUrl) throws IOException {

        Document doc = Jsoup.connect(url).get();
        Elements spanElements = doc.getElementsByAttributeValue("class", "item");
        spanElements.forEach(span -> {
            String attr = span.getElementsByTag("a").attr("href");
            String text = span.getElementsByTag("a").text();
            if (text != "" && url != "") {
                map.put(text, "https://www.kinopoisk.ru" + attr);
                map.remove("", "https://www.kinopoisk.ru");
            }
        });
        return map;
    }

    public static File createFile(String name) {

        try {
            String path = new File("").getAbsolutePath();
            File newFile = new File(path + name);
            if (newFile.createNewFile()) {
                System.out.print("Файл создан, ");
                System.out.println( "можно найти по ссылке: " + newFile.getAbsolutePath());
            } else {
                System.out.print("Файл уже существует, ");
                System.out.println( "можно найти по ссылке: " + newFile.getAbsolutePath());
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
        return new File(new File("").getAbsolutePath() + name);
    }

    public static void writeIntoExcel(File file, Map<String, String> map){

        Workbook book = new HSSFWorkbook();
        Sheet sheet = book.createSheet("NewFilm");
        int rownum = 0;
        int cellnum = 0;

        Set<Map.Entry<String, String>> set = map.entrySet();

        for (Map.Entry<String, String> entry : set) {
            Row row = sheet.createRow(rownum++);
            Cell cell = row.createCell(cellnum);
            cell.setCellValue(entry.getKey() + ": " + entry.getValue());

        }

        try {
            FileOutputStream out = new FileOutputStream(file);
            book.write(out);
            out.close();
            System.out.println("Данные успешно записаны.");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}







