package nash.tech.com.EditExcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Random;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.hssf.usermodel.HSSFDataValidation;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWriter {

	private Workbook workbook;
	private SXSSFWorkbook wb;
	private Logger logger = Logger.getLogger(ExcelWriter.class);
	private FileOutputStream fileOut;

	/**
	 * Method to create a workbook to work with excel
	 *
	 * @param filePathName
	 *            ThuanNHT
	 */
	public void createWorkBook(String filePathName) {
		if (filePathName.endsWith(".xls") || filePathName.endsWith(".XLS")) {
			workbook = new HSSFWorkbook();
		} else if (filePathName.endsWith(".xlsx") || filePathName.endsWith(".XLSX")) {
			workbook = new XSSFWorkbook();
		}
	}

	/**
	 * Method to create a new excel(xls,xlsx) file with file Name
	 *
	 * @param fileName
	 *            ThuanNHT
	 */
	public void saveToFileExcel(String filePathName) {
		try {
			fileOut = new FileOutputStream(filePathName);
			workbook.write(fileOut);
		} catch (Exception ex) {
			logger.error(ex);
		} finally {
			try {
				fileOut.close();
				workbook = null;
			} catch (IOException ex) {
				logger.error(ex);
			}
		}
	}

	public void saveToFileExcelSXSSF(String filePathName) {
		try {
			fileOut = new FileOutputStream(filePathName);
			wb.write(fileOut);
		} catch (Exception ex) {
			logger.error(ex);
		} finally {
			try {
				fileOut.close();
				wb.dispose();
				// wb = null;
			} catch (IOException ex) {
				logger.error(ex);
			}
		}
	}

	/**
	 * method to create a sheet
	 *
	 * @param sheetName
	 *            ThuanNHT
	 */
	public Sheet createSheet(String sheetName) {
		String temp = WorkbookUtil.createSafeSheetName(sheetName);
		return workbook.createSheet(temp);
	}

	/**
	 * method t create a row
	 *
	 * @param r
	 * @return ThuanNHT
	 */
	public Row createRow(Sheet sheet, int r) {
		Row row = sheet.createRow(r);
		return row;
	}

	/**
	 * method to create a cell with value
	 *
	 * @param cellValue
	 *            ThuanNHT
	 */
	public Cell createCell(Row row, int column, String cellValue) {
		// Create a cell and put a value in it.
		Cell cell = row.createCell(column);
		cell.setCellValue(cellValue);
		return cell;
	}

	/**
	 * method to create a cell with value
	 *
	 * @param cellValue
	 *            ThuanNHT
	 */
	public Cell createCell(Sheet sheet, int c, int r, String cellValue) {
		Row row = sheet.getRow(r);
		if (row == null) {
			row = sheet.createRow(r);
		}
		// Create a cell and put a value in it.
		Cell cell = row.createCell(c);
		cell.setCellValue(cellValue);
		return cell;
	}

	public Cell createCell1(Sheet sheet, int c, int r, double cellValue) {
		Row row = sheet.getRow(r);
		if (row == null) {
			row = sheet.createRow(r);
		}
		// Create a cell and put a value in it.
		Cell cell = row.createCell(c);
		cell.setCellValue(cellValue);
		return cell;
	}

	/**
	 * method to create a cell with value with style
	 *
	 * @param cellValue
	 *            ThuanNHT
	 */
	public Cell createCell(Sheet sheet, int c, int r, String cellValue, CellStyle style) {
		Row row = sheet.getRow(r);
		if (row == null) {
			row = sheet.createRow(r);
		}
		// Create a cell and put a value in it.
		Cell cell = row.createCell(c);
		cell.setCellValue(cellValue);
		cell.setCellStyle(style);
		return cell;
	}

	/**
	 * Method get primitive content Of cell
	 *
	 * @param sheet
	 * @param c
	 * @param r
	 * @return
	 */
	public static Object getCellContent(Sheet sheet, int c, int r) {
		Cell cell = getCellOfSheet(r, c, sheet);
		if (cell == null) {
			return "";
		}
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_STRING:
			return cell.getRichStringCellValue().getString();
		case Cell.CELL_TYPE_NUMERIC:
			if (DateUtil.isCellDateFormatted(cell)) {
				return cell.getDateCellValue();
			} else {
				return cell.getNumericCellValue();
			}
		case Cell.CELL_TYPE_BOOLEAN:
			return cell.getBooleanCellValue();
		case Cell.CELL_TYPE_FORMULA:
			return cell.getCellFormula();
		default:
			return "";

		}
	}

	/**
	 * Method set sheet is selected when is opened
	 *
	 * @param posSheet
	 */
	public void setSheetSelected(int posSheet) {
		try {
			workbook.setActiveSheet(posSheet);
		} catch (IllegalArgumentException ex) {
			workbook.setActiveSheet(0);
		}
	}

	public void setSheetSelectedSXSSF(int posSheet) {
		try {
			wb.setActiveSheet(posSheet);
		} catch (IllegalArgumentException ex) {
			wb.setActiveSheet(0);
		}
	}

	/**
	 * method to merge cell
	 *
	 * @param sheet
	 * @param firstRow
	 *            based 0
	 * @param lastRow
	 *            based 0
	 * @param firstCol
	 *            based 0
	 * @param lastCol
	 *            based 0
	 */
	public static void mergeCells(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
		sheet.addMergedRegion(new CellRangeAddress(firstRow, // first row
																// (0-based)
				lastRow, // last row (0-based)
				firstCol, // first column (0-based)
				lastCol // last column (0-based)
		));
	}

	/**
	 * method to fill color background for cell
	 *
	 * @param cell
	 * @param colors:BLACK,
	 *            WHITE, RED, BRIGHT_GREEN, BLUE, YELLOW, PINK, TURQUOISE,
	 *            DARK_RED, GREEN, DARK_BLUE, DARK_YELLOW, VIOLET, TEAL,
	 *            GREY_25_PERCENT, GREY_50_PERCENT, CORNFLOWER_BLUE, MAROON,
	 *            LEMON_CHIFFON, ORCHID, CORAL, ROYAL_BLUE,
	 *            LIGHT_CORNFLOWER_BLUE, SKY_BLUE, LIGHT_TURQUOISE, LIGHT_GREEN,
	 *            LIGHT_YELLOW, PALE_BLUE, ROSE, LAVENDER, TAN, LIGHT_BLUE,
	 *            AQUA, LIME, GOLD, LIGHT_ORANGE, ORANGE, BLUE_GREY,
	 *            GREY_40_PERCENT, DARK_TEAL, SEA_GREEN, DARK_GREEN,
	 *            OLIVE_GREEN, BROWN, PLUM, INDIGO, GREY_80_PERCENT, AUTOMATIC;
	 */
	public void fillAndColorCell(Cell cell, IndexedColors colors) {
		CellStyle style = workbook.createCellStyle();
		style.setFillBackgroundColor(colors.getIndex());
		cell.setCellStyle(style);
	}
	// datpk lay object tu Row

	public static Object getCellContentRow(int c, Row row) {
		Cell cell = getCellOfSheetRow(c, row);
		if (cell == null) {
			return "";
		}
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_STRING:
			return cell.getRichStringCellValue().getString();
		case Cell.CELL_TYPE_NUMERIC:
			if (DateUtil.isCellDateFormatted(cell)) {
				return cell.getDateCellValue();
			} else {
				return cell.getNumericCellValue();
			}
		case Cell.CELL_TYPE_BOOLEAN:
			return cell.getBooleanCellValue();
		case Cell.CELL_TYPE_FORMULA:
			return cell.getCellFormula();
		default:
			return "";

		}
	}

	/**
	 * Method get text content Of cell
	 *
	 * @param sheet
	 * @param c
	 * @param r
	 * @return
	 */
	public static String getCellStrContent(Sheet sheet, int c, int r) {
		Cell cell = getCellOfSheet(r, c, sheet);
		if (cell == null) {
			return "";
		}
		String temp = getCellContent(sheet, c, r).toString().trim();
		if (temp.endsWith(".0")) {
			return temp.substring(0, temp.length() - 2);
		}
		return temp;
	}

	public static String getCellStrContentString(Sheet sheet, int c, int r) {
		Cell cell = getCellOfSheet(r, c, sheet);
		if (cell == null) {
			return "";
		}
		String temp = getCellContent(sheet, c, r).toString().trim();
		return temp;
	}
	// datpk getStringconten tu Row

	public static String getCellStrContentRow(int c, Row row) {
		Cell cell = getCellOfSheetRow(c, row);
		if (cell == null) {
			return "";
		}
		String temp = getCellContentRow(c, row).toString().trim();
		if (temp.endsWith(".0")) {
			return temp.substring(0, temp.length() - 2);
		}
		return temp;
	}

	/**
	 * method to create validation from array String.But String do not exceed
	 * 255 characters
	 *
	 * @param arrValidate
	 *            * ThuanNHT
	 */
	public void createDropDownlistValidateFromArr(Sheet sheet, String[] arrValidate, int firstRow, int lastRow,
			int firstCol, int lastCol) {
		CellRangeAddressList addressList = new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol);
		DVConstraint dvConstraint = DVConstraint.createExplicitListConstraint(arrValidate);
		HSSFDataValidation dataValidation = new HSSFDataValidation(addressList, dvConstraint);
		dataValidation.setSuppressDropDownArrow(false);
		HSSFSheet sh = (HSSFSheet) sheet;
		sh.addValidationData(dataValidation);
	}

	/**
	 * Method to create validation from spread sheet via range
	 *
	 * @param range
	 * @param firstRow
	 * @param lastRow
	 * @param firstCol
	 * @param lastCol
	 *            * ThuanNHT
	 */
	public void createDropDownListValidateFromSpreadSheet(String range, int firstRow, int lastRow, int firstCol,
			int lastCol, Sheet shet) {
		Name namedRange = workbook.createName();
		Random rd = new Random();
		String refName = ("List" + rd.nextInt()).toString().replace("-", "");
		namedRange.setNameName(refName);
		// namedRange.setRefersToFormula("'Sheet1'!$A$1:$A$3");
		namedRange.setRefersToFormula(range);
		DVConstraint dvConstraint = DVConstraint.createFormulaListConstraint(refName);
		CellRangeAddressList addressList = new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol);
		HSSFDataValidation dataValidation = new HSSFDataValidation(addressList, dvConstraint);
		dataValidation.setSuppressDropDownArrow(false);
		HSSFSheet sh = (HSSFSheet) shet;
		sh.addValidationData(dataValidation);
	}

	public void createDropDownListValidateFromSpreadSheet(String sheetName, String columnRangeName, int rowRangeStart,
			int rowRangeEnd, int firstRow, int lastRow, int firstCol, int lastCol, Sheet shet) {
		String range = "'" + sheetName + "'!$" + columnRangeName + "$" + rowRangeStart + ":" + "$" + columnRangeName
				+ "$" + rowRangeEnd;
		Name namedRange = workbook.createName();
		Random rd = new Random();
		String refName = ("List" + rd.nextInt()).toString().replace("-", "");
		namedRange.setNameName(refName);
		// namedRange.setRefersToFormula("'Sheet1'!$A$1:$A$3");
		namedRange.setRefersToFormula(range);
		DVConstraint dvConstraint = DVConstraint.createFormulaListConstraint(refName);
		CellRangeAddressList addressList = new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol);
		HSSFDataValidation dataValidation = new HSSFDataValidation(addressList, dvConstraint);
		dataValidation.setSuppressDropDownArrow(false);
		HSSFSheet sh = (HSSFSheet) shet;
		sh.addValidationData(dataValidation);
	}

	public Sheet getSheetAt(int pos) {
		return workbook.getSheetAt(pos);
	}

	public Sheet getSheet(String name) {
		return workbook.getSheet(name);
	}

	/**
	 * Method to read an excel file
	 *
	 * @param filePathName
	 * @return * ThuanNHT
	 */
	public Workbook readFileExcel(String filePathName) {
		InputStream inp = null;
		try {
			inp = new FileInputStream(filePathName);
			workbook = WorkbookFactory.create(inp);
		} catch (FileNotFoundException ex) {
			logger.error(ex);
		} catch (Exception ex) {
			logger.error(ex);
		} finally {
			try {
				inp.close();
			} catch (IOException ex) {
				logger.error(ex);
			}
		}
		return workbook;
	}

	public Workbook readFileExcel(File file) {
		InputStream inp = null;
		try {
			inp = new FileInputStream(file);
			workbook = WorkbookFactory.create(inp);
		} catch (FileNotFoundException ex) {
			logger.error(ex);
		} catch (Exception ex) {
			logger.error(ex);
		} finally {
			try {
				inp.close();
			} catch (IOException ex) {
				logger.error(ex);
			}
		}
		return workbook;
	}

	public SXSSFWorkbook readFileExcelSXSSF(String filePathName) {
		InputStream inp = null;
		try {
			inp = new FileInputStream(filePathName);
			workbook = WorkbookFactory.create(inp);
			wb = new SXSSFWorkbook((XSSFWorkbook) workbook);
		} catch (FileNotFoundException ex) {
			logger.error(ex);
		} catch (Exception ex) {
			logger.error(ex);
		} finally {
			try {
				inp.close();
			} catch (IOException ex) {
				logger.error(ex);
			}
		}
		return wb;
	}

	/**
	 * * ThuanNHT
	 *
	 * @param r
	 * @param c
	 * @param sheet
	 * @return
	 */
	public static Cell getCellOfSheet(int r, int c, Sheet sheet) {
		try {
			Row row = sheet.getRow(r);
			if (row == null) {
				return null;
			}
			return row.getCell(c);
		} catch (Exception ex) {
			return null;
		}
	}

	/**
	 * set style for cell
	 *
	 * @param cell
	 * @param halign
	 * @param valign
	 * @param border
	 * @param borderColor
	 */
	public void setCellStyle(Cell cell, short halign, short valign, short border, short borderColor, int fontHeight) {
		CellStyle style = workbook.createCellStyle();
		Font font = workbook.createFont();
		font.setFontHeightInPoints((short) fontHeight);
		font.setFontName("Times New Roman");
		style.setAlignment(halign);
		style.setVerticalAlignment(valign);
		style.setBorderBottom(border);
		style.setBottomBorderColor(borderColor);
		style.setBorderLeft(border);
		style.setLeftBorderColor(borderColor);
		style.setBorderRight(border);
		style.setRightBorderColor(borderColor);
		style.setBorderTop(border);
		style.setTopBorderColor(borderColor);
		style.setFont(font);
		cell.setCellStyle(style);
	}

	public void setStandardCellStyle(Cell cell) {
		setCellStyle(cell, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER, CellStyle.BORDER_THIN,
				IndexedColors.BLACK.getIndex(), 12);
	}

	// datpk: lay cell tu Row
	public static Cell getCellOfSheetRow(int c, Row row) {
		try {
			if (row == null) {
				return null;
			}
			return row.getCell(c);
		} catch (Exception ex) {
			return null;
		}
	}

	public static Boolean compareToLong(String str, Long t) {
		Boolean check = false;
		try {
			Double d = Double.valueOf(str);
			Long l = d.longValue();
			if (l.equals(t)) {
				check = true;
			}
		} catch (Exception ex) {
			check = false;
		}
		return check;
	}

	public static Boolean doubleIsLong(String str) {
		Boolean check = false;
		try {
			Double d = Double.valueOf(str);
			Long l = d.longValue();
			if (d.equals(Double.valueOf(l))) {
				check = true;
			}
		} catch (Exception ex) {
			check = false;
		}
		return check;
	}

	public static void main(String[] arg) {
		try {
			Runtime r = Runtime.getRuntime();
			// long freeMem = r.freeMemory();
			// System.out.println("free memory before creating array: " +
			// freeMem);
			//
			ExcelWriter ewu = new ExcelWriter();
			// ewu.readFileExcel("C:\\DIEMLOM1111111111111111111.xls");
			// Sheet shet = ewu.getSheetAt(0);
			// String a = ExcelWriterUtils.getCellStrContent(shet, 8, 17);
			// ewu.createSheet("Toi la ai3");
			// for (int i = 0; i < 60000; i++) {
			// for (int j = 0; j < 4; j++) {
			// String str =
			// ExcelWriterUtils.getCellStrContent(ewu.getSheetAt(0), 2, 4);
			// Double db = Double.parseDouble(str);
			// ExcelWriterUtils.getCellStrContent(ewu.getSheetAt(0), 3, 4);
			// }
			//
			// }
			// ewu.setSheetSelected(5);
			// mergeCells(shet, 1, 5, 1, 5);
			// ewu.setStandardCellStyle(getCellOfSheet(2, 2, shet));
			// ewu.setCellStyleBottomBorder(getCellOfSheet(2, 2, shet),
			// CellStyle.BORDER_DOUBLE, IndexedColors.RED.getIndex());
			// ewu.setCellStyleBottomBorder(getCellOfSheet(2, 2, shet),
			// CellStyle.BORDER_DOUBLE, IndexedColors.GREEN.getIndex());
			// String[] arr = new String[lst.size()];
			// arr = lst.toArray(arr);
			// ewu.createDropDownlistValidation(arr);
			// ewu.readFileExcel("C:\\a.xls");
			// System.out.println(ewu.getCellStrContent(ewu.getSheetAt(0), 1, 1)
			// + "------------");
			// System.out.println(ewu.getCellContent(ewu.getSheetAt(0), 1, 1) +
			// "------------");
			// ewu.setSheet(ewu.getWorkbook().getSheetAt(0));
			// ewu.createDropDownListValidateFromSpreadSheet("'Sheet2'!$B$4:$B$11",
			// 8, 12, 4, 5, ewu.getWorkbook().getSheetAt(0));
			// Cell cel = getCellOfSheet(17, 5, sh1);
			// ewu.saveToFileExcel("C:\\a.xlsx");

			// freeMem = r.freeMemory();
			// System.out.println("free memory after creating array: " +
			// freeMem);
			// r.gc();
			// freeMem = r.freeMemory();
			// System.out.println("free memory after running gc(): " + freeMem);
			//
			// ExcelWriterUtils ewu2 = new ExcelWriterUtils();
			// ewu2.readFileExcel("C:\\aaa.xls");
			//// Sheet shet2 = ewu2.createSheet("Toi la ai3");
			// for (int i = 0; i < 60000; i++) {
			// for (int j = 0; j < 4; j++) {
			// ewu2.createCell(ewu2.getSheetAt(1), j, i, "Hang thu " + i);
			// }
			//
			// }
			Workbook wb = new SXSSFWorkbook();

			FileInputStream inp = new FileInputStream("D:\\ReportWave.xlsx");
			Workbook workbook = WorkbookFactory.create(inp);
			Sheet tmpSheet = workbook.getSheetAt(0);
			Row row = tmpSheet.getRow(3);
			Cell cellDate = ExcelWriter.getCellOfSheet(4, 0, tmpSheet);
			CellStyle csCellDate = cellDate.getCellStyle();

			// Font ff = csCellDate.getFontIndex();
			System.out.println(row.getHeight());
			System.out.println(csCellDate.getFillForegroundColorColor());
			System.out.println(csCellDate.getFillForegroundColor());
			System.out.println(csCellDate.getFontIndex());
			System.out.println(csCellDate.getAlignment());
			System.out.println(csCellDate.getVerticalAlignment());

			wb = new SXSSFWorkbook((XSSFWorkbook) workbook);
			ewu.setWb((SXSSFWorkbook) wb);
			Cell c = null;
			CellStyle cs = wb.createCellStyle();
			cs.setFillForegroundColor((short) 44);
			cs.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
			cs.setBorderBottom((short) 1);
			cs.setBorderLeft((short) 1);
			cs.setBorderRight((short) 1);
			cs.setBorderTop((short) 1);
			cs.setWrapText(true);
			cs.setAlignment((short) 2);
			cs.setVerticalAlignment((short) 1);
			Font f = wb.createFont();
			f.setBoldweight(Font.BOLDWEIGHT_BOLD);
			f.setFontHeightInPoints((short) 11);
			f.setFontName("Times New Roman");
			cs.setFont(f);
			Sheet she = ewu.getSheetAtSXSSF(0);
			//
			//
			for (int i = 9; i < 20; i++) {
				for (int j = 0; j < 14; j++) {
					ewu.createCell(she, j, i, "cell[ " + i + "," + j + "]").setCellStyle(cs);
					she.getRow(i).setHeight((short) 600);

				}

			}
			//
			//// String[] arr = new String[lst.size()];
			//// arr = lst.toArray(arr);
			//// ewu.createDropDownlistValidation(arr);
			//// ewu.readFileExcel("C:\\a.xls");
			//// System.out.println(ewu.getCellStrContent(ewu.getSheetAt(0), 1,
			// 1) + "------------");
			//// System.out.println(ewu.getCellContent(ewu.getSheetAt(0), 1, 1)
			// + "------------");
			//// ewu.setSheet(ewu.getWorkbook().getSheetAt(0));
			//// ewu.createDropDownListValidateFromSpreadSheet("'Sheet2'!$B$4:$B$11",
			// 8, 12, 4, 5, ewu.getWorkbook().getSheetAt(0));
			//// Cell cel = getCellOfSheet(17, 5, sh1);

			ewu.saveToFileExcelSXSSF("D:\\aaaa.xlsx");
		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}

	public Workbook getWorkbook() {
		return workbook;
	}

	public void setWorkbook(Workbook workbook) {
		this.workbook = workbook;
	}

	public SXSSFWorkbook getWb() {
		return wb;
	}

	public void setWb(SXSSFWorkbook wb) {
		this.wb = wb;
	}

	// R3490.1_trangltt9_new_25022013_start
	public void saveToFile(OutputStream out) {
		try {
			workbook.write(out);
		} catch (Exception ex) {
			logger.error(ex);
		} finally {
			workbook = null;
		}
	}
	// R3490.1_trangltt9_new_25022013_end

	public Sheet getSheetAtSXSSF(int pos) {
		return wb.getSheetAt(pos);
	}
}