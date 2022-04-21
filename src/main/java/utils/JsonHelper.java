package utils;

import model.NewFilm;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.select.Elements;

import java.io.IOException;
import java.util.List;

public class JsonHelper {

   private final static String mainUrl = "https://www.kinopoisk.ru";

    public static List<NewFilm> parseUrl(String requestUrl, List<NewFilm> newFilms) throws IOException {
        Document doc = Jsoup.connect(requestUrl).get();
        Elements spanElements = doc.getElementsByAttributeValue("class", "item");
        spanElements.forEach(span -> {
            String attr = span.getElementsByTag("a").attr("href");
            String text = span.getElementsByTag("a").text();
            if (text != "" && requestUrl != "") {
                String link = mainUrl + attr;
                newFilms.add( new NewFilm(text, link));
            }
        });
        return newFilms;
    }
}
