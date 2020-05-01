import GoogleSheetUtils.GoogleSheetReader;
import GoogleSheetUtils.GoogleSheetWriter;

import java.io.IOException;
import java.security.GeneralSecurityException;
import java.util.List;

public class Main {
    public static void main(String[] args) throws IOException, GeneralSecurityException {
        GoogleSheetReader googleSheetReader = new GoogleSheetReader();
        List<List<Object>> response = googleSheetReader.getSheetContent("Letters", "A1", "C1");

        for(List row: response) {
            for(Object item: row) {
                System.out.println(item);
            }
        }
    }
}
