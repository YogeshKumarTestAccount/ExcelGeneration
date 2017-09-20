package com.charter.excelapi;

/**
 *@author YOGESH KUMAR
 *@ykumar10
 *03/07/2017
 *
 */

import java.util.List;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class CharterExcelCreation {

	public static void main(String args[]) {

		CharterExcelCreation ch = new CharterExcelCreation();

		/*
		 * ch.charterValidationSheet("D:\\", "1445550901143934",
		 * "1220901143910", "8245120901143910", "NPDC_SUSPEND", "Inter",
		 * "1220901143910_8245120901143910_000007", "1010000011be", "",
		 * "UIM PASS");
		 */

		/*
		 * ch.charterValidationSheet("D:\\", "1445550901143934",
		 * "1220901143910", "8245120901143910", "NPDC_RESUME", "Inter",
		 * "1220901143910_8245120901143910_000006", "1010000011f1", "",
		 * "UIM PASS");
		 */

		/*
		 * ch.charterValidationSheet("D:\\", "1445550901143934",
		 * "1220901143910", "8245120901143910", "NPDC_RESUME", "Inter",
		 * "1220901143910_8245120901143910_000007", "1010000011be", "",
		 * "MacLogging PASS");
		 */

		ch.charterValidationSheet("D:\\", "1445550901143934", "1220901143910",
				"8245120901143910", "NPDC1_RESUME2", "Inter",
				"1220901143910_8245120901143910_000006", "1010000011f1", "",
				"BACC FAIL");

	}

	public void charterValidationSheet(String jenkinsLogDirPath,
			String csgOrder, String customerId, String accounId,
			String scenario, String servicetype, String ServiceId,
			String cmmac, String mtamac, String validationType) {
		boolean isCustomerIdExist = false;
		boolean isServiceTypeExist = false;
		boolean isScenarioExist = false;
		boolean isServiceIdExist = false;
		boolean checkValidation = true;
		String checkRowExistence = null;
		HSSFWorkbook workbook = null;
		Workbook workbookFectory = null;
		HSSFSheet sheet = null;
		Row headerRow = null;
		Row startContentRow = null;
		FileOutputStream outFile = null;
		int fixColumn = 7;
		Cell cell = null;
		String FileNameAsCurrentDate = new SimpleDateFormat(
				"'_'dd-MM-yyyy'.xls'").format(new Date());
		String fileName = jenkinsLogDirPath
				+ customerId.concat(FileNameAsCurrentDate);
		File isFileExist = new File(fileName);

		// Check the File already Exist execute if block otherwise execute else
		// block
		try {
			if (isFileExist.exists() && isFileExist.isFile()) {
				// boolean checkRowExistence=true;
				// List<String> listForRowExistenceCheck = ArrayList<String> ();

				// Workbook is Advance interface which can work both .xls or
				// .xlsx
				// File Format.

				FileInputStream file = new FileInputStream(isFileExist);
				workbookFectory = WorkbookFactory.create(file);
				Sheet updateSheet = workbookFectory.getSheet("Sheet0");
				headerRow = updateSheet.getRow(0);
				int activationClmCount = headerRow.getLastCellNum() - 8;
				int updateRowcounter = 0;
				int breakLoopifRowFound = 0;
				// int howManyRow=0;
				Iterator<Row> rowIterator = updateSheet.iterator();
				while (rowIterator.hasNext()) {

					Row row = rowIterator.next();
					if (checkRowExistence != null) {

						if (checkRowExistence.equals("Not Row Found")) {
							updateRowcounter = updateRowcounter + 1;
						}
					}

					if (breakLoopifRowFound == 0) {
						isCustomerIdExist = false;
						isServiceTypeExist = false;
						isScenarioExist = false;
						isServiceIdExist = false;

						Iterator<Cell> cellIterator = row.cellIterator();
						while (cellIterator.hasNext()) {
							Cell duplicateCellCheacker = cellIterator.next();

							if (duplicateCellCheacker.getStringCellValue()
									.equals(customerId)) {
								isCustomerIdExist = true;
							}

							if (duplicateCellCheacker.getStringCellValue()
									.equals(scenario)) {

								isScenarioExist = true;
							}
							if (duplicateCellCheacker.getStringCellValue()
									.equals(servicetype)) {
								isServiceTypeExist = true;
							}

							if (duplicateCellCheacker.getStringCellValue()
									.equals(ServiceId)) {
								isServiceIdExist = true;
							}

						}
					}
					if (isCustomerIdExist && isScenarioExist
							&& isServiceTypeExist && isServiceIdExist) {
						checkRowExistence = "Row Found";
						breakLoopifRowFound = 1;

					} else {
						checkRowExistence = "Not Row Found";

					}

				}

				// If Row Already exist (checkRowExistence=true) then Update
				// Require column at existing row.
				if (checkRowExistence.equals("Row Found")) {
					startContentRow = updateSheet.getRow(updateRowcounter);

					for (int j = 0; j < headerRow.getLastCellNum(); j++) {

						// Update CMMMAC
						if ((headerRow.getCell(j).getStringCellValue())
								.equals("CMMAC")) {
							// Create a new cell in current row
							cell = startContentRow.createCell(j);
							// Set value for new cell value
							if (cmmac != null) {
								cell.setCellValue(cmmac);
							}
						}
						// Update MTAMAC
						if ((headerRow.getCell(j).getStringCellValue())
								.equals("MTAMAC")) {

							cell = startContentRow.createCell(j);
							// Set value for new cell value
							if (mtamac != null) {
								cell.setCellValue(mtamac);
							}

						}

						// Update the Validation Type Header and its value if
						// Not
						// Exist
						if (validationType != null && checkValidation
								&& checkValidation && j > fixColumn) {
							HSSFCellStyle style = (HSSFCellStyle) workbookFectory
									.createCellStyle();
							HSSFFont font = (HSSFFont) workbookFectory
									.createFont();
							font.setFontName(HSSFFont.FONT_ARIAL);
							font.setFontHeightInPoints((short) 13);
							font.setBold(true);
							style.setFont(font);
							if ((headerRow.getCell(j).getStringCellValue())
									.equals(calculateValidationHeader(validationType)
											+ " Validation")) {
								cell = startContentRow.createCell(j);
								cell.setCellValue(calculateValidationValue(validationType));
								checkValidation = false;
							} else {
								// creating new column by Increasing header by
								// +1
								if (activationClmCount == 1) {

									if (headerRow
											.getCell(
													headerRow.getLastCellNum() - 1)
											.getStringCellValue()
											.equals("Overall Status")) {

										cell = headerRow.createCell(headerRow
												.getLastCellNum() - 1);
										cell.setCellValue(calculateValidationHeader(validationType)
												+ " Validation");
										cell.setCellStyle(style);

										// Auto Size Specific Row
										HSSFRow row = (HSSFRow) workbookFectory
												.getSheetAt(0).getRow(0);
										for (int colNum = 0; colNum < row
												.getLastCellNum(); colNum++) {

											workbookFectory.getSheetAt(0)
													.autoSizeColumn(colNum);
										}
										// Set Value (FIX for Overriding
										// OverallStatus value)
										cell = startContentRow
												.createCell(headerRow
														.getLastCellNum() - 1);
										cell.setCellValue(calculateValidationValue(validationType));

									} else {
										cell = headerRow.createCell(headerRow
												.getLastCellNum());
										// Set New Column Header Name
										cell.setCellValue(calculateValidationHeader(validationType)
												+ " Validation");

										cell.setCellStyle(style);
										// Auto Size Specific Row
										HSSFRow row = (HSSFRow) workbookFectory
												.getSheetAt(0).getRow(0);
										for (int colNum = 0; colNum < row
												.getLastCellNum(); colNum++) {
											workbookFectory.getSheetAt(0)
													.autoSizeColumn(colNum);
										}

										// Set Value
										cell = startContentRow
												.createCell(headerRow
														.getLastCellNum() - 1);
										cell.setCellValue(calculateValidationValue(validationType));

									}

									checkValidation = false;
								}
								activationClmCount = activationClmCount - 1;
							}

						}
						updateSheet.autoSizeColumn(j);
						file.close();

					}
				} else {

					// Creating New Content Row inside the New Sheet
					createNewRowInsideSheet(updateSheet, csgOrder, customerId,
							accounId, scenario, servicetype, ServiceId, cmmac,
							mtamac, (HSSFWorkbook) workbookFectory,
							validationType);

					file.close();

				}
			} else {
				// if Worksheet not exist create new it with Header and value

				workbook = new HSSFWorkbook();
				sheet = workbook.createSheet("Sheet0");

				// Creating New Header Row inside the New Sheet
				createSheetHeader(sheet, cell, headerRow, workbook);

				// Creating New Content Row inside the New Sheet
				createNewRow(sheet, csgOrder, customerId, accounId, scenario,
						servicetype, ServiceId, cmmac, mtamac, workbook,
						validationType);

			}

			outFile = new FileOutputStream(new File(fileName));
			if (workbook != null) {
				workbook.write(outFile);

			}
			if (workbookFectory != null) {
				workbookFectory.write(outFile);
			}

		} catch (FileNotFoundException e1) {

			e1.printStackTrace();
		} catch (IOException e2) {
			e2.printStackTrace();
		}

		catch (EncryptedDocumentException e3) {
			e3.printStackTrace();
		}

		catch (InvalidFormatException e4) {
			e4.printStackTrace();
		}

		finally {
			try {
				outFile.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}

		writeOverallTestStatus(fileName);
	}

	// Creating Sheet Header

	public void createSheetHeader(Sheet sheet, Cell cell, Row headerRow,
			HSSFWorkbook workbook) {

		headerRow = sheet.createRow(0);
		int initval;
		HSSFCellStyle style = workbook.createCellStyle();
		HSSFFont font = workbook.createFont();
		font.setFontName(HSSFFont.FONT_ARIAL);
		font.setFontHeightInPoints((short) 13);
		font.setBold(true);
		style.setFont(font);
		for (initval = 0; initval <= OSMHeaderValues.values().length - 1; initval++) {

			cell = headerRow.createCell(initval);
			cell.setCellStyle(style);
			boolean newColumnFlag = true;
			if (OSMHeaderValues.values()[initval].getName().equals(
					"CSG Order ID")) {
				cell.setCellValue(OSMHeaderValues.values()[initval].getName());
				newColumnFlag = false;
			}

			if (OSMHeaderValues.values()[initval].getName().equals(
					"Customer ID")) {
				cell.setCellValue(OSMHeaderValues.values()[initval].getName());
				newColumnFlag = false;
			}
			if (OSMHeaderValues.values()[initval].getName()
					.equals("Account ID")) {
				cell.setCellValue(OSMHeaderValues.values()[initval].getName());
				newColumnFlag = false;
			}

			if (OSMHeaderValues.values()[initval].getName().equals("Scenario")) {
				cell.setCellValue(OSMHeaderValues.values()[initval].getName());
				newColumnFlag = false;
			}

			if (OSMHeaderValues.values()[initval].getName().equals(
					"Service Type")) {
				cell.setCellValue(OSMHeaderValues.values()[initval].getName());
				newColumnFlag = false;
			}

			if (OSMHeaderValues.values()[initval].getName()
					.equals("Service ID")) {
				cell.setCellValue(OSMHeaderValues.values()[initval].getName());
				newColumnFlag = false;
			}

			if (OSMHeaderValues.values()[initval].getName().equals("CMMAC")) {
				cell.setCellValue(OSMHeaderValues.values()[initval].getName());
				newColumnFlag = false;
			}

			if (OSMHeaderValues.values()[initval].getName().equals("MTAMAC")) {
				cell.setCellValue(OSMHeaderValues.values()[initval].getName());
				newColumnFlag = false;
			}

			if (OSMHeaderValues.values()[initval].getName().equals(
					"UIM Validation")) {
				cell.setCellValue(OSMHeaderValues.values()[initval].getName());
				newColumnFlag = false;
			}

			if (OSMHeaderValues.values()[initval].getName().equals(
					"PRO Validatoin")) {
				cell.setCellValue(OSMHeaderValues.values()[initval].getName());
				newColumnFlag = false;
			}

			if (newColumnFlag) {
				cell.setCellValue(OSMHeaderValues.values()[initval].getName());

			}
			sheet.autoSizeColumn(initval);

		}

	}

	// Creating New Row for New WorkBook
	public void createNewRow(Sheet sheet, String csgOrder, String customerId,
			String accounId, String scenario, String serviceType,
			String ServiceId, String cmmac, String mtamac,
			HSSFWorkbook hssfWorkbook, String validationType) {

		// Create a new row in current sheet

		int countRow = sheet.getLastRowNum() + 1;
		boolean checkValidation = true;
		Row newRow = sheet.createRow(countRow);
		Row headerRow = sheet.getRow(0);
		int countCol = headerRow.getLastCellNum();
		HSSFCellStyle style = hssfWorkbook.createCellStyle();
		HSSFFont font = hssfWorkbook.createFont();
		font.setFontName(HSSFFont.FONT_ARIAL);
		font.setFontHeightInPoints((short) 13);
		font.setBold(true);
		style.setFont(font);
		Cell cell = null;
		for (int l = 0; l < countCol; l++) {
			if ((headerRow.getCell(l).getStringCellValue())
					.equals("CSG Order ID")) {

				// Create a new cell in current row
				cell = newRow.createCell(l);
				// Set value for new cell
				cell.setCellValue(csgOrder);

			}

			if ((headerRow.getCell(l).getStringCellValue())
					.equals("Customer ID")) {
				// Create a new cell in current row
				cell = newRow.createCell(l);
				// Set value for new cell
				cell.setCellValue(customerId);
			}

			if ((headerRow.getCell(l).getStringCellValue())
					.equals("Account ID")) {
				// Create a new cell in current row
				cell = newRow.createCell(l);
				// Set value for new cell
				cell.setCellValue(accounId);
			}
			if ((headerRow.getCell(l).getStringCellValue()).equals("Scenario")) {
				// Create a new cell in current row
				cell = newRow.createCell(l);
				// Set value for new cell
				cell.setCellValue(scenario);
			}
			if ((headerRow.getCell(l).getStringCellValue())
					.equals("Service Type")) {
				// Create a new cell in current row
				cell = newRow.createCell(l);
				// Set value for new cell
				cell.setCellValue(serviceType);
			}
			if ((headerRow.getCell(l).getStringCellValue())
					.equals("Service ID")) {
				// Create a new cell in current row
				cell = newRow.createCell(l);
				// Set value for new cell
				cell.setCellValue(ServiceId);
			}
			if ((headerRow.getCell(l).getStringCellValue()).equals("CMMAC")) {
				// Create a new cell in current row
				cell = newRow.createCell(l);
				// Set value for new cell
				cell.setCellValue(cmmac);
			}
			if ((headerRow.getCell(l).getStringCellValue()).equals("MTAMAC")) {
				// Create a new cell in current row
				cell = newRow.createCell(l);
				// Set value for new cell
				cell.setCellValue(mtamac);

			}

			if (validationType != null && checkValidation) {
				cell = headerRow.createCell(headerRow.getLastCellNum());
				// Set New Column Header Name
				cell.setCellValue(calculateValidationHeader(validationType)
						+ " Validation");
				cell.setCellStyle(style);
				// Auto Size Specific Row
				HSSFRow row = hssfWorkbook.getSheetAt(0).getRow(0);
				for (int colNum = 0; colNum < row.getLastCellNum(); colNum++) {

					hssfWorkbook.getSheetAt(0).autoSizeColumn(colNum);
				}

				// Set Value
				cell = newRow.createCell(headerRow.getLastCellNum() - 1);
				cell.setCellValue(calculateValidationValue(validationType));
				checkValidation = false;

			}
			sheet.autoSizeColumn(l);

		}

	}

	// Creating New Row inside the Sheet for Existing WorkBook
	public void createNewRowInsideSheet(Sheet sheet, String csgOrder,
			String customerId, String accounId, String scenario,
			String serviceType, String ServiceId, String cmmac, String mtamac,
			HSSFWorkbook hssfWorkbook, String validationType) {

		// Create a new row in current sheet

		int countRow = sheet.getLastRowNum() + 1;
		boolean checkValidation = true;
		int fixColumn = 7;
		Row newRow = sheet.createRow(countRow);
		Row headerRow = sheet.getRow(0);
		int countCol = headerRow.getLastCellNum();
		HSSFCellStyle style = hssfWorkbook.createCellStyle();
		HSSFFont font = hssfWorkbook.createFont();
		font.setFontName(HSSFFont.FONT_ARIAL);
		font.setFontHeightInPoints((short) 13);
		font.setBold(true);
		style.setFont(font);
		Cell cell = null;
		int activationClmCount = countCol - 8;
		for (int l = 0; l < countCol; l++) {
			if ((headerRow.getCell(l).getStringCellValue())
					.equals("CSG Order ID")) {

				// Create a new cell in current row
				cell = newRow.createCell(l);
				// Set value for new cell
				cell.setCellValue(csgOrder);

			}

			if ((headerRow.getCell(l).getStringCellValue())
					.equals("Customer ID")) {
				// Create a new cell in current row
				cell = newRow.createCell(l);
				// Set value for new cell
				cell.setCellValue(customerId);
			}

			if ((headerRow.getCell(l).getStringCellValue())
					.equals("Account ID")) {
				// Create a new cell in current row
				cell = newRow.createCell(l);
				// Set value for new cell
				cell.setCellValue(accounId);
			}
			if ((headerRow.getCell(l).getStringCellValue()).equals("Scenario")) {
				// Create a new cell in current row
				cell = newRow.createCell(l);
				// Set value for new cell
				cell.setCellValue(scenario);
			}
			if ((headerRow.getCell(l).getStringCellValue())
					.equals("Service Type")) {
				// Create a new cell in current row
				cell = newRow.createCell(l);
				// Set value for new cell
				cell.setCellValue(serviceType);
			}
			if ((headerRow.getCell(l).getStringCellValue())
					.equals("Service ID")) {
				// Create a new cell in current row
				cell = newRow.createCell(l);
				// Set value for new cell
				cell.setCellValue(ServiceId);
			}
			if ((headerRow.getCell(l).getStringCellValue()).equals("CMMAC")) {
				// Create a new cell in current row
				cell = newRow.createCell(l);
				// Set value for new cell
				cell.setCellValue(cmmac);
			}
			if ((headerRow.getCell(l).getStringCellValue()).equals("MTAMAC")) {
				// Create a new cell in current row
				cell = newRow.createCell(l);
				// Set value for new cell
				cell.setCellValue(mtamac);

			}

			if (validationType != null && checkValidation && l > fixColumn) {
				// Check header Already Exist put the value .if not then create
				// the New column header and its value
				if ((headerRow.getCell(l).getStringCellValue())
						.equals(calculateValidationHeader(validationType)
								+ " Validation")) {
					cell = newRow.createCell(l);
					cell.setCellValue(calculateValidationValue(validationType));
					checkValidation = false;
				} else {
					// creating new column by Increasing header by +1

					if (activationClmCount == 1) {
						cell = headerRow.createCell(headerRow.getLastCellNum());
						// Set New Column Header Name
						cell.setCellValue(calculateValidationHeader(validationType)
								+ " Validation");
						cell.setCellStyle(style);
						// Auto Size Specific Row
						HSSFRow row = hssfWorkbook.getSheetAt(0).getRow(0);
						for (int colNum = 0; colNum < row.getLastCellNum(); colNum++) {

							hssfWorkbook.getSheetAt(0).autoSizeColumn(colNum);
						}

						// Set Value
						cell = newRow
								.createCell(headerRow.getLastCellNum() - 1);
						cell.setCellValue(calculateValidationValue(validationType));
						checkValidation = false;
					}

				}

				activationClmCount = activationClmCount - 1;
			}
			sheet.autoSizeColumn(l);

		}

	}

	public static String calculateValidationValue(String stringValue) {
		// String s1=s.substring(0,s.length()-4);
		// String s2=s.substring(s.length()-4, s.length());
		String dyValue = null;
		char str[] = stringValue.toCharArray();
		for (int i = stringValue.length() - 1; i > 0; i--) {
			if (str[i] == ' ') {
				dyValue = stringValue.substring(i + 1, stringValue.length());
				break;
			}
		}
		return dyValue;
	}

	public static String calculateValidationHeader(String stringHeader) {
		String dyHeader = null;
		char str[] = stringHeader.toCharArray();
		for (int i = stringHeader.length() - 1; i > 0; i--) {
			if (str[i] == ' ') {
				dyHeader = stringHeader.substring(0, i);
				break;
			}
		}
		return dyHeader;
	}

	// Method Write The OverAll Test Status
	public void writeOverallTestStatus(String fileName) {

		File file = new File(fileName);
		Workbook workbookFectory = null;
		Row headerRow = null;
		Cell cell = null;
		try {
			FileInputStream fis = new FileInputStream(file);
			workbookFectory = WorkbookFactory.create(fis);
			Sheet updateSheet = workbookFectory.getSheet("Sheet0");
			headerRow = updateSheet.getRow(0);
			int activationClmCount = headerRow.getLastCellNum() - 9;
			HSSFCellStyle style = (HSSFCellStyle) workbookFectory
					.createCellStyle();
			HSSFFont font = (HSSFFont) workbookFectory.createFont();
			font.setFontName(HSSFFont.FONT_ARIAL);
			font.setFontHeightInPoints((short) 13);
			font.setBold(true);
			style.setFont(font);
			int j = 0;
			if (activationClmCount > 0)

			{

				if (!(headerRow.getCell(headerRow.getLastCellNum() - 1)
						.getStringCellValue().equals("Overall Status"))) {
					cell = headerRow.createCell(headerRow.getLastCellNum());
					cell.setCellValue("Overall Status");
					cell.setCellStyle(style);
					// Auto Size Specific Row
					cell.setCellStyle(style);
					HSSFRow row = (HSSFRow) workbookFectory.getSheetAt(0)
							.getRow(0);
					for (int colNum = 0; colNum < row.getLastCellNum(); colNum++) {

						workbookFectory.getSheetAt(0).autoSizeColumn(colNum);
					}

				}

				for (int i = 1; i <= updateSheet.getLastRowNum(); i++) {
					Row startContentRow = updateSheet.getRow(i);
					boolean flag = true;
					for (j = 8; j <= startContentRow.getLastCellNum() - 1; j++) {
						if (startContentRow.getCell(j) != null) {
							String checkOverAllStatus = startContentRow
									.getCell(j).getStringCellValue();
							if (checkOverAllStatus.equals(("FAIL"))) {
								cell = startContentRow.createCell(headerRow
										.getLastCellNum() - 1);
								cell.setCellValue("FAIL");
								flag = false;
								break;

							}
						}

					}

					if (flag) {
						if (startContentRow.getCell(j) != null

								&& startContentRow.getCell(j - 1)
										.getStringCellValue().equals("PASS")) {

							cell = startContentRow.createCell(headerRow
									.getLastCellNum() - 1);
							cell.setCellValue("PASS");
						} else {

							cell = startContentRow.createCell(headerRow
									.getLastCellNum() - 1);
							cell.setCellValue("PASS");

						}

					}

					updateSheet.autoSizeColumn(j - 1);
					fis.close();
					FileOutputStream outFile = new FileOutputStream(new File(
							fileName));
					workbookFectory.write(outFile);
					outFile.close();

				}
			}
		} catch (FileNotFoundException e1) {
			e1.printStackTrace();
		} catch (IOException e2) {
			e2.printStackTrace();
		}

		catch (EncryptedDocumentException e3) {
			e3.printStackTrace();
		}

		catch (InvalidFormatException e4) {
			e4.printStackTrace();
		}

	}

}