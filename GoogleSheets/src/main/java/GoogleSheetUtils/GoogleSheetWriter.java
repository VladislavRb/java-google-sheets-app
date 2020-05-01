package GoogleSheetUtils;

import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.UpdateValuesResponse;
import com.google.api.services.sheets.v4.model.ValueRange;
import com.google.api.services.sheets.v4.model.AppendValuesResponse;

import java.io.IOException;
import java.security.GeneralSecurityException;
import java.util.Arrays;
import java.util.List;

public class GoogleSheetWriter {
    private Sheets sheetsService;
    private String SPREADSHEET_ID = "1XhBUJ2o946YUXzrPGf6kMnEjc4-RK1V7kzKYJBhmLug";
    private GoogleSheetServiceProvider googleSheetServiceProvider = new GoogleSheetServiceProvider();

    public void insertRow(String sheetName, List<Object> row) throws IOException, GeneralSecurityException {
        this.sheetsService = this.googleSheetServiceProvider.getSheetsService();
        ValueRange appendBody = new ValueRange().setValues(Arrays.asList(row));
        AppendValuesResponse appendResult = this.sheetsService.spreadsheets().values()
                .append(SPREADSHEET_ID, sheetName, appendBody)
                .setValueInputOption("USER_ENTERED")
                .setInsertDataOption("INSERT_ROWS")
                .setIncludeValuesInResponse(true)
                .execute();
    }

    public void updateCell(String cellName, Object newValue) throws IOException, GeneralSecurityException {
        this.sheetsService = this.googleSheetServiceProvider.getSheetsService();
        ValueRange updateBody = new ValueRange().setValues(Arrays.asList(Arrays.asList(newValue)));
        UpdateValuesResponse updateResult = this.sheetsService.spreadsheets().values()
                .update(SPREADSHEET_ID, cellName, updateBody)
                .setValueInputOption("RAW")
                .execute();
    }
}
