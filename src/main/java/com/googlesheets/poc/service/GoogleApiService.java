package com.googlesheets.poc.service;


import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.springframework.stereotype.Service;

import com.google.api.client.googleapis.auth.oauth2.GoogleCredential;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.CellData;
import com.google.api.services.sheets.v4.model.GridData;
import com.google.api.services.sheets.v4.model.RowData;
import com.google.api.services.sheets.v4.model.Sheet;
import com.google.api.services.sheets.v4.model.Spreadsheet;
import com.googlesheets.poc.model.GRow;
import com.googlesheets.poc.model.GSheet;

@Service
public class GoogleApiService{

	private static final String APPLICATION_NAME ="Google Sheets API Java GooGLeSheetService";


    private static FileDataStoreFactory DATA_STORE_FACTORY;

    private static final com.google.api.client.json.JsonFactory JSON_FACTORY =
        JacksonFactory.getDefaultInstance();

    private static HttpTransport HTTP_TRANSPORT;

    private static final List<String> SCOPES = Arrays.asList(SheetsScopes.SPREADSHEETS_READONLY);

    static {
        try {
            HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
        } catch (Throwable t) {
            t.printStackTrace();
        }
    }

    /**
     * Build and return an authorized Sheets API client service.
     * @return an authorized Sheets API client service
     * @throws IOException
     */
    private Sheets getSheetsService( String token) throws IOException {
    	GoogleCredential credential = new GoogleCredential().setAccessToken(token);

        return new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, credential)
                .setApplicationName(APPLICATION_NAME)
                .build();
    }
    public static void main(String[] args) {
		try {
			new GoogleApiService().findSheet(null,null);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

    public List<GSheet> findSheet(String spreadSheetId, String token) throws IOException {
    	
    	List<GSheet> sheetRes = new ArrayList<GSheet>();
        // Build a new authorized API client service.
        Sheets sheets = getSheetsService(token);
        Sheets.Spreadsheets.Get request = sheets.spreadsheets().get(spreadSheetId);
        request.setIncludeGridData(true);
        Spreadsheet response = request.execute();
        if (response == null) {
            System.out.println("No data found.");
        } else {
			System.out.println("Name, Major" + response);
			for (Sheet sheet : response.getSheets()) {
				if (sheet != null && sheet.getData() != null) {
					GSheet sheetData = new GSheet();
					sheetRes.add(sheetData);
					sheetData.setName(sheet.getProperties().getTitle());
					System.out.println("---Sheet:" + sheet.getProperties().getTitle());
					for (GridData grid : sheet.getData()) {
						if (grid != null && grid.getRowData() != null) {
							for (RowData row : grid.getRowData()) {
								GRow gRow = new GRow();
								sheetData.getRows().add(gRow);
								System.out.println();
								if (row != null && row.getValues() != null) {
									for (CellData cell : row.getValues()) {
										if (cell != null && cell.getEffectiveValue() != null) {
											Object cellVal =  cell.getEffectiveValue().get(cell.getEffectiveValue().keySet().iterator().next());
											if (cellVal != null) {
												System.out.print(cellVal + ":");
												gRow.getColumns().add(cellVal.toString());
											}
										}
									}
								}
							}
						}
					}
				}
				System.out.println();
			}
        }
        return sheetRes;
    }


}