package nash.tech.com.EditExcel;

import java.io.File;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Hello world!
 *
 */
public class App {
	public static void main(String[] args) {
		String folderPath = "D:/TEST";
		App app = new App();
		app.readFolder(folderPath);
	}

	public void readFolder(String folderPath) {
		File folder = new File(folderPath);
		File[] listOfFiles = folder.listFiles();

		for (int i = 0; i < listOfFiles.length; i++) {
			if (listOfFiles[i].isFile()) {
				System.out.println("File " + listOfFiles[i].getAbsolutePath());
				editFile(listOfFiles[i].getAbsolutePath());
			} else if (listOfFiles[i].isDirectory()) {
				System.out.println("Directory " + listOfFiles[i].getAbsolutePath());
			}
		}
	}

	public void editFile(String filePath) {
		ExcelWriter excelWriter = new ExcelWriter();
		Workbook workbook = excelWriter.readFileExcel(filePath);
		Sheet sheet3 = workbook.getSheetAt(3);
		Cell cell = ExcelWriter.getCellOfSheet(4, 0, sheet3);
		CellStyle cellStyle = cell.getCellStyle();
		cellStyle.setWrapText(true);
		String content = "画面リンク \n (Link màn hình)";
		Sheet sheet2 = workbook.getSheetAt(2);
		excelWriter.createCell(sheet3, 0, 5, content, cellStyle);
		excelWriter.createCell(sheet2, 0, 6, content, cellStyle);
		excelWriter.setWorkbook(workbook);
		excelWriter.saveToFileExcel(filePath);
		System.out.println("FINISH");
	}
}
