package web.excel;

import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

//Helper Class
public class ExcelConverter {
	private MultipartFile file;
	
	ExcelConverter(){
		file = null;
	}
	
	ExcelConverter(MultipartFile f){
		file = f;
	}
	
	
	//Nutzen die Apache POI API um die Excel DAtei zu lesen und die Ergebnisse in eine neue Datei zu schreiben
	public Workbook excel() throws EncryptedDocumentException, IOException {
		Workbook wb = WorkbookFactory.create(file.getInputStream());
		Workbook newWb = new XSSFWorkbook();

		Sheet sh = wb.getSheetAt(0);
		Sheet newSh = newWb.createSheet("Ergebnisse");
		
		int rowNumber = sh.getLastRowNum();
		
		//durchlaufen alle Reihen der Datei
		for (int i = 0; i <= rowNumber; i++) {
			Row newRow = newSh.createRow(i);
			Row tmpRow = sh.getRow(i);
			
			Cell newCell = newRow.createCell(0);
			
			Cell firstCell = tmpRow.getCell(0);
			Cell operatorCell = tmpRow.getCell(1);
			Cell thirdCell = tmpRow.getCell(2);
			
			//kontrollieren Format
			if(!((firstCell == null || operatorCell ==null) || thirdCell == null)) {
				if((firstCell.getCellType() == CellType.NUMERIC && thirdCell.getCellType() == CellType.NUMERIC) && operatorCell.getCellType() == CellType.STRING) {
					Double firstValue = firstCell.getNumericCellValue();
					String operator = operatorCell.getStringCellValue();
					Double secondValue = thirdCell.getNumericCellValue();
					
					newCell.setCellValue(doMath(firstValue, secondValue, operator));
				}
				else {
					newCell.setCellValue("Falsches Format");
				}
			}
			else {
				newCell.setCellValue("Falsches Format");
			}
		}
		
		return newWb;
	}
	
	//kontrollieren welcher Operator genutzt wird und geben das Ergebniss zurück
	public double doMath(double firstOperand, double secondOperand, String operator) {
		switch(operator) {
			case "+" : return firstOperand + secondOperand; 
			case "-" : return firstOperand - secondOperand;
			case "/" : return firstOperand / secondOperand; 
			//falls durch 0 geteilt wird wird in der Excel Datei in der entsprechenden Zelle ein Fehler angegeben
			case "*" : return firstOperand * secondOperand;
			default  : return 0; 
		}
	}
}
