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

    private static final String BRACKET_LEFT = "(";
    private static final String BRACKET_RIGHT = ")";
    private static final String LOAD_ACCOUNTS_SH = "./load_accounts.sh ";
    private static final String NEW_LINE_SYMBOL = "\n";
    private static final String XLSX_FILE_PATH = "test.xlsx";
    private static final String SPACE = " ";
    private static int dateRowNumber;

    public static void main(String[] args) throws IOException, InvalidFormatException {

	Workbook workbook = WorkbookFactory.create(new File(XLSX_FILE_PATH));
	Sheet sheet = workbook.getSheetAt(0);
	DataFormatter dataFormatter = new DataFormatter();
	SimpleDateFormat sdf = new SimpleDateFormat();
	Date currentDate = new Date();

	getDateContainingRowNum(sheet, currentDate, sdf, dataFormatter);
	// System.out.println("daterownumber is " + dateRowNumber);
	removeRowsBeforeCurrentDate(sheet);

	List<String> accountsIdsForMerchantDataProd = new ArrayList<String>();
	List<String> accountsIdsForMerchantDataStage = new ArrayList<String>();
	List<String> accountsIdsForSubidProd = new ArrayList<String>();
	List<String> accountsIdsForSubidStage = new ArrayList<String>();

	fillInAccountIdsList(sheet, dataFormatter, ReportTypes.MerchantData.name(), accountsIdsForMerchantDataProd,
		ServerTypes.PROD);
	fillInAccountIdsList(sheet, dataFormatter, ReportTypes.MerchantData.name(), accountsIdsForMerchantDataStage,
		ServerTypes.STAGE);
	fillInAccountIdsList(sheet, dataFormatter, ReportTypes.Subid.name(), accountsIdsForSubidProd, ServerTypes.PROD);
	fillInAccountIdsList(sheet, dataFormatter, ReportTypes.Subid.name(), accountsIdsForSubidStage,
		ServerTypes.STAGE);

	printShScript(accountsIdsForMerchantDataProd, currentDate, sdf, true, ServerTypes.PROD);
	printShScript(accountsIdsForMerchantDataStage, currentDate, sdf, true, ServerTypes.STAGE);
	printShScript(accountsIdsForSubidProd, currentDate, sdf, false, ServerTypes.PROD);
	printShScript(accountsIdsForSubidStage, currentDate, sdf, false, ServerTypes.STAGE);
	printMigrateScript(accountsIdsForMerchantDataProd, ServerTypes.PROD);
	printMigrateScript(accountsIdsForMerchantDataStage, ServerTypes.STAGE);
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
	    List<String> accountsIdsList, ServerTypes type) {
	sheet.forEach(row -> {
	    Cell reportTypeCell = row.getCell(10);
	    Cell accountIdCell = row.getCell(5);
	    String reportTypeCellValue = dataFormatter.formatCellValue(reportTypeCell);
	    String accountIdCellValue = dataFormatter.formatCellValue(accountIdCell);
	    if (reportTypeCellValue.equals(reportType)) {
		if (!accountIdCellValue.contains(NEW_LINE_SYMBOL)) {
		    if (accountIdCellValue.contains(BRACKET_LEFT)) {
			accountIdCellValue = filterString(accountIdCellValue, type);
		    }
		    accountsIdsList.add(accountIdCellValue);
		} else {
		    accountsIdsList.addAll(removeNewLines(accountIdCellValue, type));
		}
	    }
	});
    }

    private static List<String> removeNewLines(String cellValue, ServerTypes type) {
	String[] cellValuesArray = cellValue.split(NEW_LINE_SYMBOL);
	List<String> tempList = new ArrayList<String>();
	for (String value : cellValuesArray) {
	    if (value.contains(BRACKET_LEFT)) {
		value = filterString(value, type);
	    }
	    tempList.add(value);
	}
	return tempList;
    }

    private static String filterString(String celValue, ServerTypes type) {
	int index_left = celValue.indexOf(BRACKET_LEFT);
	int index_right = celValue.indexOf(BRACKET_RIGHT);

	switch (type) {
	case PROD:
	    celValue = celValue.substring(0, index_left - 1);
	    break;
	case STAGE:
	    celValue = celValue.substring(index_left + 1, index_right);
	}
	return celValue;
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
	    boolean isMerchantData, ServerTypes type) {
	String reportType = (isMerchantData) ? ReportTypes.MerchantData.name() : ReportTypes.Subid.name();
	String idsAsSpaceSeparatedString = String.join(SPACE, accountsIdsList);
	idsAsSpaceSeparatedString.trim();

	System.out.println(new StringBuilder("Script for ").append(reportType).append(" on ").append(type).append(":"));
	System.out.println(new StringBuilder(LOAD_ACCOUNTS_SH).append(reportType).append(SPACE)
		.append(getDateForScript(currentDate, sdf)).append(SPACE).append(idsAsSpaceSeparatedString));
    }

    private static void printMigrateScript(List<String> accountsIdsList, ServerTypes type) {
	String idsAsSpaceSeparatedString = String.join(",", accountsIdsList);
	System.out.println(new StringBuilder("Migration script for ").append(ReportTypes.MerchantData).append(" on ").append(type)
		.append(":"));
	System.out.println("{");
	System.out.println(new StringBuilder("    \"reportType\":\"").append(ReportTypes.MerchantData).append("\","));
	System.out.println(new StringBuilder("    \"accID\":\"").append(idsAsSpaceSeparatedString).append("\""));
	System.out.println("}");
    }
}