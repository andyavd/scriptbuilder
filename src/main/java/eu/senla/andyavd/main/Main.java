package eu.senla.andyavd.main;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

public class Main {

    private static final String LOAD_ACCOUNTS_SH = "./load_accounts.sh ";
    private static final String NEW_LINE_SYMBOL = "\n";
    private static final String XLSX_FILE_PATH = "test.xlsx";
    private static final String SPACE = " ";
    private static int dateRowNumber;

    public static void main(String[] args) throws IOException, InvalidFormatException {

	// RowIndex for Account ids = 5
	// RowIndex for ReportType = 10

	Workbook workbook = WorkbookFactory.create(new File(XLSX_FILE_PATH));
	Sheet sheet = workbook.getSheetAt(0);
	DataFormatter dataFormatter = new DataFormatter();
	SimpleDateFormat sdf = new SimpleDateFormat();
	Date currentDate = new Date();

	getDateContainingRowNum(sheet, currentDate, sdf, dataFormatter);
	// System.out.println("daterownumber is " + dateRowNumber);
	removeRowsBeforeCurrentDate(sheet);

	List<String> accountsIdsForMerchantData = new ArrayList<String>();
	List<String> accountsIdsForSubid = new ArrayList<String>();

	fillInAccountIdsList(sheet, dataFormatter, ReportTypes.MerchantData.name(), accountsIdsForMerchantData);
	fillInAccountIdsList(sheet, dataFormatter, ReportTypes.Subid.name(), accountsIdsForSubid);

	printShScript(accountsIdsForMerchantData, currentDate, sdf, true);
	printShScript(accountsIdsForSubid, currentDate, sdf, false);

	// Closing the workbook
	workbook.close();

    }

    private static void getDateContainingRowNum(Sheet sheet, Date currentDate, SimpleDateFormat sdf,
	    DataFormatter dataFormatter) {
	sdf.applyPattern("M/d/yy");
	String date = sdf.format(currentDate);
	sheet.forEach(row -> {
	    row.forEach(cell -> {
		String cellValue = dataFormatter.formatCellValue(cell);
		if (cellValue.equals(date)) {
		    dateRowNumber = cell.getRowIndex();
		}
	    });
	});
    }

    private static void removeRowsBeforeCurrentDate(Sheet sheet) {
	for (int i = 0; i <= dateRowNumber; i++) {
	    Row row = sheet.getRow(i);
	    sheet.removeRow(row);
	}
    }

    private static void fillInAccountIdsList(Sheet sheet, DataFormatter dataFormatter, String reportType,
	    List<String> accountsIdsList) {
	sheet.forEach(row -> {
	    Cell reportTypeCell = row.getCell(10);
	    Cell accountIdCell = row.getCell(5);
	    String reportTypeCellValue = dataFormatter.formatCellValue(reportTypeCell);
	    String accountIdCellValue = dataFormatter.formatCellValue(accountIdCell);
	    if (reportTypeCellValue.equals(reportType)) {
		if (accountIdCellValue.contains(NEW_LINE_SYMBOL)) {
		    accountIdCellValue = accountIdCellValue.replace(NEW_LINE_SYMBOL, SPACE);
		}
		accountsIdsList.add(accountIdCellValue);
	    }
	});
    }

    private static String getDateForScript(Date currentDate, SimpleDateFormat sdf) {
	sdf.applyPattern("dd/MM/yyyy");
	Calendar calendar = Calendar.getInstance();
	calendar.setTime(currentDate);
	calendar.set(Calendar.DAY_OF_MONTH, calendar.getActualMinimum(Calendar.DAY_OF_MONTH));
	StringBuilder sb = new StringBuilder(sdf.format(calendar.getTime())).append(SPACE)
		.append(sdf.format(currentDate));
	return sb.toString();
    }

    private static void printShScript(List<String> accountsIdsList, Date currentDate, SimpleDateFormat sdf,
	    boolean isMerchantData) {
	String reportType = (isMerchantData) ? ReportTypes.MerchantData.name() : ReportTypes.Subid.name();
	String idsAsSpaceSeparatedString = String.join(SPACE, accountsIdsList);
	idsAsSpaceSeparatedString.trim();

	System.out.println(new StringBuilder("Script for ").append(reportType).append(":"));
	System.out.println(new StringBuilder(LOAD_ACCOUNTS_SH).append(reportType).append(SPACE)
		.append(getDateForScript(currentDate, sdf)).append(SPACE).append(idsAsSpaceSeparatedString));
    }
}