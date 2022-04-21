import model.NewFilm;
import utils.FileHelper;
import utils.JsonHelper;
import java.io.File;
import java.io.IOException;
import java.util.*;

public class Parser {

    public static void main(String[] args) throws IOException {
        String url = "https://www.kinopoisk.ru/premiere/ru/company/4/";
        List<NewFilm> list = new ArrayList<>();
        List<NewFilm> newFilms = JsonHelper.parseUrl(url, list);
        if (newFilms.size()>0) {
            System.out.println("Список премьер с сайта " + "\"kinopoisk\"" + " получен.");
            Date date = new Date();
            System.out.println("Текущая дата: " + date);
            System.out.println("Для вывода нажмите:");
            System.out.println("\"1\" для вывода в КОНСОЛЬ");
            System.out.println("\"2\" для вывода в MS EXEL");
            getScanner(newFilms);
        } else {
            System.out.println("ОШИБКА...\nСписок премьер с сайта " + "\"kinopoisk\"" + " не получен.");

        }
    }

    private static void getScanner(List<NewFilm> newFilms) {
        Scanner i = new Scanner(System.in);
        int num = i.nextInt();
        switch (num) {
            case 1:
                for (NewFilm newFilm : newFilms) {
                    System.out.println(newFilm.getName() + ": " + newFilm.getLink());
                }
                break;
            case 2:
                File file = FileHelper.createFile("\\model.NewFilm.xls");
                boolean isSuccessWrite = FileHelper.writeExcel(file, newFilms);
                if (isSuccessWrite) {
                    System.out.println("Данные успешно записаны.");
                } else {
                    System.out.println("При записи возникла ошибка...");
                }
                break;
            default:
                System.out.println("Некорректный ввод!");
                System.out.println("Программа завершена.");
        }
        i.close();
    }
}







