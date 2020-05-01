package GoogleSheetUtils;

import java.io.IOException;
import java.security.GeneralSecurityException;
import java.util.List;

import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.ValueRange;

public class GoogleSheetReader {
    private Sheets sheetsService;
    private String SPREADSHEET_ID = "1XhBUJ2o946YUXzrPGf6kMnEjc4-RK1V7kzKYJBhmLug";
    private GoogleSheetServiceProvider googleSheetServiceProvider = new GoogleSheetServiceProvider();

    public List<List<Object>> getSheetContent(String sheetName, String upperLeftCellName, String lowerRightCellName) throws IOException, GeneralSecurityException {
        this.sheetsService = this.googleSheetServiceProvider.getSheetsService();
        String range = "" + sheetName + "!" + upperLeftCellName + ":" + lowerRightCellName;

        ValueRange response = this.sheetsService.spreadsheets().values()
                .get(SPREADSHEET_ID, range)
                .execute();

        return response.getValues();
    }
}

