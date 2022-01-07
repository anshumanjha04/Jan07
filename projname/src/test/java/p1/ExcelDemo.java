package p1;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelDemo {

	public static void main(String[] args) throws EncryptedDocumentException, FileNotFoundException, IOException {
		
		String path = "./data/book.xlsx";
		//Open the Excel
		Workbook wb = WorkbookFactory.create(new FileInputStream(path));
		//Read the data
		String v = wb.getSheet("sheet1").getRow(0).getCell(0).getStringCellValue();
		System.out.println(v);
		
		//Write the data
		wb.getSheet("sheet1").getRow(0).getCell(0).setCellValue("Anshuman");
		wb.write(new FileOutputStream(path));
		//Close the Excel file
		wb.close();

	}

}
