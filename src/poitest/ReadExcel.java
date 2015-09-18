package poitest;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.usermodel.*;

public class ReadExcel {
	public static void main(String[] args) throws FileNotFoundException, IOException {
		// Will contain cell name / value pair for input cells and output cells
		Map<String, String> inputCellsMap = new HashMap<String, String>();
		Map<String, String> outputCellsMap = new HashMap<String, String>();
		
		// Open the Excel file
		FileInputStream file = new FileInputStream(new File(args[1]));
		
		// Get the current workbook
		HSSFWorkbook workbook = new HSSFWorkbook(file);
		
		// Get the first sheet of the workbook
		HSSFSheet sheet = workbook.getSheetAt(0);
		
		// Get the input cells that need to be modified and
		// store their name and value in the inputCellsMap
		for (String element : args[3].split(";")) {
			inputCellsMap.put(element.split("=")[0], element.split("=")[1]);
		}
		
		// Loop through the cells that need to be modified and 
		// set the new value in the Excel document
		Iterator<Entry<String,String>> iterator = inputCellsMap.entrySet().iterator();
		while (iterator.hasNext()) {
			Map.Entry<String,String> entry = (Map.Entry<String,String>) iterator.next();

			CellReference cellReferenceInput = new CellReference(entry.getKey());
			int cellReferenceInputRow = cellReferenceInput.getRow();
			int cellReferenceInputColumn = cellReferenceInput.getCol();

			Row rowInput = sheet.getRow(cellReferenceInputRow);
			if (rowInput == null)
			    rowInput = sheet.createRow(cellReferenceInputRow);
			Cell cell = rowInput.getCell(cellReferenceInputColumn, Row.CREATE_NULL_AS_BLANK);				
			cell.setCellValue(Integer.parseInt(entry.getValue()));		
		}
					
		// Apply all formulas after altering cell values
		HSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);

		// Get result
		// This is currently hard coded
		// Will have to get results from the cells entered as args
		Cell resultCell = sheet.getRow(6).getCell(1);				

		System.out.println(resultCell.getNumericCellValue());	
						
		workbook.close();		
	}
}