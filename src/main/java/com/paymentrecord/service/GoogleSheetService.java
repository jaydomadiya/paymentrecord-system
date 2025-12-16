package com.paymentrecord.service;

import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.*;
import com.paymentrecord.config.GoogleSheetConfig;
import com.paymentrecord.dto.PaymentRequest;
import org.springframework.stereotype.Service;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;

@Service
public class GoogleSheetService {

    private static final String SPREADSHEET_ID = "1WBGxzx8Tx-z1YcXTJYka9dkaSLlvBrTvoSXQKmn664g";
    private static final int TOTAL_COLUMNS = 26; // A to Z columns
    private static final int SUMMARY_START_ROW = 100; // Summary starts at row 100

    private final DateTimeFormatter dateFormatter = DateTimeFormatter.ofPattern("dd-MM-yyyy");

    public void savePayment(PaymentRequest req) throws Exception {
        Sheets sheets = GoogleSheetConfig.getSheetsService();

        // Get current month's sheet name
        String currentMonthSheet = getOrCreateMonthlySheet(sheets, req.getDate());

        // Save payment in current month's sheet
        savePaymentToSheet(sheets, req, currentMonthSheet);

        // Update daily total automatically
        updateDailyTotal(sheets, currentMonthSheet, req.getDate());
    }

    private String getOrCreateMonthlySheet(Sheets sheets, LocalDate date) throws Exception {
        String sheetName = date.format(DateTimeFormatter.ofPattern("MMM-yyyy"));

        // Get all sheets
        Spreadsheet spreadsheet = sheets.spreadsheets().get(SPREADSHEET_ID).execute();
        List<Sheet> existingSheets = spreadsheet.getSheets();

        // Check if sheet exists
        boolean sheetExists = existingSheets.stream()
                .anyMatch(sheet -> sheet.getProperties().getTitle().equals(sheetName));

        // Create new sheet if doesn't exist OR fix old sheet
        if (!sheetExists) {
            createNewMonthSheet(sheets, sheetName, date);
        } else {
            // Fix old sheet if it has only 8 columns
            fixOldSheetGrid(sheets, sheetName);
        }

        return sheetName;
    }

    private void fixOldSheetGrid(Sheets sheets, String sheetName) throws Exception {
        // Get sheet ID first
        Integer sheetId = getSheetId(sheets, sheetName);
        if (sheetId == null) return;

        // Check current grid size
        Spreadsheet spreadsheet = sheets.spreadsheets().get(SPREADSHEET_ID).execute();
        Sheet targetSheet = spreadsheet.getSheets().stream()
                .filter(s -> s.getProperties().getTitle().equals(sheetName))
                .findFirst()
                .orElse(null);

        if (targetSheet != null) {
            GridProperties gridProps = targetSheet.getProperties().getGridProperties();

            // If sheet has only 8 columns, expand it to 26
            if (gridProps.getColumnCount() != null && gridProps.getColumnCount() < TOTAL_COLUMNS) {
                List<Request> requests = new ArrayList<>();

                requests.add(new Request().setUpdateSheetProperties(
                        new UpdateSheetPropertiesRequest()
                                .setProperties(new SheetProperties()
                                        .setSheetId(sheetId)
                                        .setGridProperties(new GridProperties()
                                                .setRowCount(1000)
                                                .setColumnCount(TOTAL_COLUMNS)))
                                .setFields("gridProperties")));

                BatchUpdateSpreadsheetRequest batchUpdate = new BatchUpdateSpreadsheetRequest()
                        .setRequests(requests);
                sheets.spreadsheets().batchUpdate(SPREADSHEET_ID, batchUpdate).execute();

                System.out.println("Fixed sheet '" + sheetName + "' grid to " + TOTAL_COLUMNS + " columns");
            }

            // Initialize summary if not exists
            initializeSummaryIfNeeded(sheets, sheetName, LocalDate.now());
        }
    }

    private void initializeSummaryIfNeeded(Sheets sheets, String sheetName, LocalDate date) throws Exception {
        // Try to get summary header - if error, it doesn't exist
        try {
            sheets.spreadsheets().values()
                    .get(SPREADSHEET_ID, sheetName + "!J" + SUMMARY_START_ROW)
                    .execute();
        } catch (Exception e) {
            // Summary doesn't exist, create it
            initializeSummarySection(sheets, sheetName, date);
        }
    }

    private void createNewMonthSheet(Sheets sheets, String sheetName, LocalDate date) throws Exception {
        List<Request> requests = new ArrayList<>();

        // Create new sheet with 26 columns
        requests.add(new Request().setAddSheet(new AddSheetRequest()
                .setProperties(new SheetProperties()
                        .setTitle(sheetName)
                        .setGridProperties(new GridProperties()
                                .setRowCount(1000)
                                .setColumnCount(TOTAL_COLUMNS)))));

        // Apply batch update
        BatchUpdateSpreadsheetRequest batchUpdateRequest = new BatchUpdateSpreadsheetRequest()
                .setRequests(requests);
        sheets.spreadsheets().batchUpdate(SPREADSHEET_ID, batchUpdateRequest).execute();

        // Initialize the new sheet
        initializeNewSheet(sheets, sheetName, date);
    }

    private void initializeNewSheet(Sheets sheets, String sheetName, LocalDate date) throws Exception {
        // Headers for main data (A-H)
        List<List<Object>> headers = Arrays.asList(
                Arrays.asList("Date", "Channel Type", "User Name", "UPI ID",
                        "Amount (₹)", "Status", "Daily Total", "Remarks")
        );

        ValueRange headerBody = new ValueRange().setValues(headers);
        sheets.spreadsheets().values()
                .update(SPREADSHEET_ID, sheetName + "!A1:H1", headerBody)
                .setValueInputOption("RAW")
                .execute();

        // Apply formatting
        applyMainHeaderFormatting(sheets, sheetName);

        // Initialize summary section
        initializeSummarySection(sheets, sheetName, date);
    }

    private void applyMainHeaderFormatting(Sheets sheets, String sheetName) throws Exception {
        Integer sheetId = getSheetId(sheets, sheetName);
        if (sheetId == null) return;

        List<Request> requests = new ArrayList<>();

        // Header formatting for A-H
        requests.add(new Request()
                .setRepeatCell(new RepeatCellRequest()
                        .setRange(new GridRange()
                                .setSheetId(sheetId)
                                .setStartRowIndex(0)
                                .setEndRowIndex(1)
                                .setStartColumnIndex(0)
                                .setEndColumnIndex(8))
                        .setCell(new CellData()
                                .setUserEnteredFormat(new CellFormat()
                                        .setBackgroundColor(new Color().setRed(0.2f).setGreen(0.4f).setBlue(0.6f))
                                        .setTextFormat(new TextFormat()
                                                .setBold(true)
                                                .setFontSize(12)
                                                .setForegroundColor(new Color().setRed(1f).setGreen(1f).setBlue(1f)))
                                        .setHorizontalAlignment("CENTER")))
                        .setFields("userEnteredFormat")));

        // Freeze header row
        requests.add(new Request()
                .setUpdateSheetProperties(new UpdateSheetPropertiesRequest()
                        .setProperties(new SheetProperties()
                                .setSheetId(sheetId)
                                .setGridProperties(new GridProperties()
                                        .setFrozenRowCount(1)))
                        .setFields("gridProperties.frozenRowCount")));

        BatchUpdateSpreadsheetRequest batchUpdateRequest = new BatchUpdateSpreadsheetRequest()
                .setRequests(requests);
        sheets.spreadsheets().batchUpdate(SPREADSHEET_ID, batchUpdateRequest).execute();
    }

    private void initializeSummarySection(Sheets sheets, String sheetName, LocalDate date) throws Exception {
        // Monthly summary data - placed at row 100
        List<List<Object>> summaryData = Arrays.asList(
                Arrays.asList("Monthly Summary for " + date.format(DateTimeFormatter.ofPattern("MMMM yyyy")), ""),
                Arrays.asList("Date", "Total Amount (₹)"),
                Arrays.asList("", ""),
                Arrays.asList("Grand Total:", "=SUM(K" + (SUMMARY_START_ROW + 3) + ":K999)")
        );

        ValueRange summaryBody = new ValueRange().setValues(summaryData);
        String summaryRange = sheetName + "!J" + SUMMARY_START_ROW + ":K" + (SUMMARY_START_ROW + 4);

        sheets.spreadsheets().values()
                .update(SPREADSHEET_ID, summaryRange, summaryBody)
                .setValueInputOption("USER_ENTERED")
                .execute();

        // Apply summary formatting
        applySummaryFormatting(sheets, sheetName);
    }

    private void applySummaryFormatting(Sheets sheets, String sheetName) throws Exception {
        Integer sheetId = getSheetId(sheets, sheetName);
        if (sheetId == null) return;

        List<Request> requests = new ArrayList<>();

        // Summary header formatting
        requests.add(new Request()
                .setRepeatCell(new RepeatCellRequest()
                        .setRange(new GridRange()
                                .setSheetId(sheetId)
                                .setStartRowIndex(SUMMARY_START_ROW - 1)
                                .setEndRowIndex(SUMMARY_START_ROW)
                                .setStartColumnIndex(9) // Column J
                                .setEndColumnIndex(11)) // Column K
                        .setCell(new CellData()
                                .setUserEnteredFormat(new CellFormat()
                                        .setBackgroundColor(new Color().setRed(0.8f).setGreen(0.4f).setBlue(0.2f))
                                        .setTextFormat(new TextFormat()
                                                .setBold(true)
                                                .setFontSize(14)
                                                .setForegroundColor(new Color().setRed(1f).setGreen(1f).setBlue(1f)))
                                        .setHorizontalAlignment("CENTER")))
                        .setFields("userEnteredFormat")));

        // Grand Total formatting
        requests.add(new Request()
                .setRepeatCell(new RepeatCellRequest()
                        .setRange(new GridRange()
                                .setSheetId(sheetId)
                                .setStartRowIndex(SUMMARY_START_ROW + 3)
                                .setEndRowIndex(SUMMARY_START_ROW + 4)
                                .setStartColumnIndex(9)
                                .setEndColumnIndex(11))
                        .setCell(new CellData()
                                .setUserEnteredFormat(new CellFormat()
                                        .setBackgroundColor(new Color().setRed(1f).setGreen(0.9f).setBlue(0f))
                                        .setTextFormat(new TextFormat().setBold(true).setFontSize(12))
                                        .setHorizontalAlignment("RIGHT")
                                        .setNumberFormat(new NumberFormat()
                                                .setType("CURRENCY")
                                                .setPattern("\"₹\"#,##0.00"))))
                        .setFields("userEnteredFormat")));

        BatchUpdateSpreadsheetRequest batchUpdateRequest = new BatchUpdateSpreadsheetRequest()
                .setRequests(requests);
        sheets.spreadsheets().batchUpdate(SPREADSHEET_ID, batchUpdateRequest).execute();
    }

    private void savePaymentToSheet(Sheets sheets, PaymentRequest req, String sheetName) throws Exception {
        // Find next empty row
        ValueRange response = sheets.spreadsheets().values()
                .get(SPREADSHEET_ID, sheetName + "!A:A")
                .execute();

        int nextRow = (response.getValues() != null) ? response.getValues().size() + 1 : 2;

        // Prepare row data
        String formattedDate = req.getDate().format(dateFormatter);
        List<Object> row = Arrays.asList(
                formattedDate,
                req.getChannelType(),
                req.getUserName(),
                req.getUpiId(),
                req.getAmount(),
                req.getStatus(),
                "", // Daily Total (will be auto-calculated)
                ""  // Remarks
        );

        ValueRange body = new ValueRange().setValues(Collections.singletonList(row));

        // Insert row
        sheets.spreadsheets().values()
                .update(SPREADSHEET_ID, sheetName + "!A" + nextRow + ":H" + nextRow, body)
                .setValueInputOption("USER_ENTERED")
                .execute();

        // Apply row formatting
        applyRowFormatting(sheets, sheetName, nextRow);
    }

    private void applyRowFormatting(Sheets sheets, String sheetName, int row) throws Exception {
        Integer sheetId = getSheetId(sheets, sheetName);
        if (sheetId == null) return;

        List<Request> requests = new ArrayList<>();

        // Alternate row coloring
        Color rowColor = (row % 2 == 0) ?
                new Color().setRed(0.95f).setGreen(0.95f).setBlue(0.95f) :
                new Color().setRed(1f).setGreen(1f).setBlue(1f);

        requests.add(new Request()
                .setRepeatCell(new RepeatCellRequest()
                        .setRange(new GridRange()
                                .setSheetId(sheetId)
                                .setStartRowIndex(row - 1)
                                .setEndRowIndex(row)
                                .setStartColumnIndex(0)
                                .setEndColumnIndex(8))
                        .setCell(new CellData()
                                .setUserEnteredFormat(new CellFormat()
                                        .setBackgroundColor(rowColor)
                                        .setBorders(new Borders()
                                                .setBottom(new Border()
                                                        .setStyle("SOLID")
                                                        .setColor(new Color().setRed(0.8f).setGreen(0.8f).setBlue(0.8f))))))
                        .setFields("userEnteredFormat(backgroundColor,borders)")));

        // Amount column formatting
        requests.add(new Request()
                .setRepeatCell(new RepeatCellRequest()
                        .setRange(new GridRange()
                                .setSheetId(sheetId)
                                .setStartRowIndex(row - 1)
                                .setEndRowIndex(row)
                                .setStartColumnIndex(4)
                                .setEndColumnIndex(5))
                        .setCell(new CellData()
                                .setUserEnteredFormat(new CellFormat()
                                        .setNumberFormat(new NumberFormat()
                                                .setType("CURRENCY")
                                                .setPattern("\"₹\"#,##0.00"))))
                        .setFields("userEnteredFormat.numberFormat")));

        BatchUpdateSpreadsheetRequest batchUpdateRequest = new BatchUpdateSpreadsheetRequest()
                .setRequests(requests);
        sheets.spreadsheets().batchUpdate(SPREADSHEET_ID, batchUpdateRequest).execute();
    }

    private void updateDailyTotal(Sheets sheets, String sheetName, LocalDate date) throws Exception {
        String targetDate = date.format(dateFormatter);

        // Get all rows for this date
        ValueRange response = sheets.spreadsheets().values()
                .get(SPREADSHEET_ID, sheetName + "!A:H")
                .execute();

        List<List<Object>> values = response.getValues();
        if (values == null || values.size() <= 1) return;

        double dailyTotal = 0;
        int lastRowForDate = -1;

        // Calculate total for the date
        for (int i = 1; i < values.size(); i++) {
            List<Object> row = values.get(i);
            if (row.size() > 0 && targetDate.equals(row.get(0))) {
                if (row.size() > 4 && row.get(4) != null && !row.get(4).toString().isEmpty()) {
                    try {
                        String amountStr = row.get(4).toString()
                                .replace("₹", "")
                                .replace(",", "")
                                .trim();
                        dailyTotal += Double.parseDouble(amountStr);
                    } catch (NumberFormatException e) {
                        // Ignore
                    }
                }
                lastRowForDate = i + 1;
            }
        }

        // Update daily total in column G
        if (lastRowForDate != -1 && dailyTotal > 0) {
            String dailyTotalCell = sheetName + "!G" + lastRowForDate;
            ValueRange totalBody = new ValueRange()
                    .setValues(Collections.singletonList(
                            Collections.singletonList(dailyTotal)
                    ));

            sheets.spreadsheets().values()
                    .update(SPREADSHEET_ID, dailyTotalCell, totalBody)
                    .setValueInputOption("USER_ENTERED")
                    .execute();

            // Update summary
            updateSummaryWithDate(sheets, sheetName, targetDate, dailyTotal);
        }
    }

    private void updateSummaryWithDate(Sheets sheets, String sheetName, String date, double total) throws Exception {
        // Get current summary data
        String summaryRange = sheetName + "!J" + (SUMMARY_START_ROW + 2) + ":K999";
        ValueRange summaryResponse = sheets.spreadsheets().values()
                .get(SPREADSHEET_ID, summaryRange)
                .execute();

        List<List<Object>> summaryData = new ArrayList<>();
        boolean dateFound = false;
        int dataRow = SUMMARY_START_ROW + 2;

        if (summaryResponse.getValues() != null) {
            for (List<Object> row : summaryResponse.getValues()) {
                if (row.size() > 0 && row.get(0) != null && !row.get(0).toString().isEmpty()) {
                    if (row.get(0).toString().equals(date)) {
                        // Update existing entry
                        summaryData.add(Arrays.asList(date, total));
                        dateFound = true;
                    } else {
                        summaryData.add(row);
                    }
                    dataRow++;
                }
            }
        }

        // Add new date if not found
        if (!dateFound) {
            summaryData.add(Arrays.asList(date, total));
        }

        // Write back summary data
        if (!summaryData.isEmpty()) {
            ValueRange updateBody = new ValueRange().setValues(summaryData);
            sheets.spreadsheets().values()
                    .update(SPREADSHEET_ID, sheetName + "!J" + (SUMMARY_START_ROW + 2), updateBody)
                    .setValueInputOption("USER_ENTERED")
                    .execute();
        }
    }

    private Integer getSheetId(Sheets sheets, String sheetName) throws Exception {
        Spreadsheet spreadsheet = sheets.spreadsheets().get(SPREADSHEET_ID).execute();
        return spreadsheet.getSheets().stream()
                .filter(sheet -> sheet.getProperties().getTitle().equals(sheetName))
                .map(sheet -> sheet.getProperties().getSheetId())
                .findFirst()
                .orElse(null);
    }
}
