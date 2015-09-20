package poitest;

import java.util.List;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.formula.eval.ErrorEval;
import org.apache.poi.ss.usermodel.*;

public class ReadExcel {
	public static void main(String[] args) throws FileNotFoundException, IOException {
		// Will contain cell name / value pair for input cells			
		Map<String, String> inputCellsMap = new HashMap<String, String>();
		
		// Will contain cell name for output cells
		List<String> outputCells = new ArrayList<String>();
		
		// Open the Excel file
		FileInputStream file = new FileInputStream(new File(args[0]));
		
		// Get the current workbook
		HSSFWorkbook workbook = new HSSFWorkbook(file);			
		
		// Get the input cells that need to be modified and
		// store their name and value in the inputCellsMap
		for (String element : args[1].split(";")) {
			inputCellsMap.put(element.split("=")[0], element.split("=")[1]);
		}
		
		// Get the output cells that will be accessed for resulting values
		for (String element : args[2].split(";")) {
			outputCells.add(element);			
		}
		
		// Loop through the cells that need to be modified and 
		// set the new value in the Excel document
		Iterator<Entry<String,String>> inputIterator = inputCellsMap.entrySet().iterator();
		while (inputIterator.hasNext()) {
			Map.Entry<String,String> inputEntry = (Map.Entry<String,String>) inputIterator.next();

			CellReference cellReferenceInput = new CellReference(inputEntry.getKey());
			int cellReferenceInputRow = cellReferenceInput.getRow();
			int cellReferenceInputColumn = cellReferenceInput.getCol();

			// Get sheet name for each input cell
			HSSFSheet inputSheet = workbook.getSheet(inputEntry.getKey().split("!")[0]);
			
			Row rowInput = inputSheet.getRow(cellReferenceInputRow);
			if (rowInput == null)
			    rowInput = inputSheet.createRow(cellReferenceInputRow);
			Cell cellInput = rowInput.getCell(cellReferenceInputColumn, Row.CREATE_NULL_AS_BLANK);				
			cellInput.setCellValue(Integer.parseInt(inputEntry.getValue()));		
		}
					
		// Apply all formulas after altering cell values		
		workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();		
		
		// Get the results from the output cells
		for (int i = 0; i < outputCells.size(); i++) {
			CellReference cellReferenceOutput = new CellReference(outputCells.get(i));
			int cellReferenceOutputRow = cellReferenceOutput.getRow();
			int cellReferenceOutputColumn = cellReferenceOutput.getCol();
			
			// Get sheet name for each output cell
			HSSFSheet outputSheet = workbook.getSheet(outputCells.get(i).split("!")[0]);
			
			Row rowOutput = outputSheet.getRow(cellReferenceOutputRow);
			Cell cellOutput = rowOutput.getCell(cellReferenceOutputColumn, Row.CREATE_NULL_AS_BLANK);
			
			// Display results
			switch (cellOutput.getCellType()) {
				case Cell.CELL_TYPE_BOOLEAN:
					System.out.println(cellOutput.getBooleanCellValue());
					break;
				case Cell.CELL_TYPE_NUMERIC:
					System.out.println(cellOutput.getNumericCellValue());
					break;
				case Cell.CELL_TYPE_STRING:
					System.out.println(cellOutput.getStringCellValue());
					break;
				case Cell.CELL_TYPE_BLANK:
					break;				
				case Cell.CELL_TYPE_FORMULA:							
					switch (cellOutput.getCachedFormulaResultType()) {
						case Cell.CELL_TYPE_STRING:
							System.out.println(cellOutput.getRichStringCellValue());							
							break;
						case Cell.CELL_TYPE_NUMERIC:
							HSSFCellStyle style = (HSSFCellStyle) cellOutput.getCellStyle();
							if (style == null) {
								System.out.println(cellOutput.getNumericCellValue());
							} else {
								DataFormatter formatter = new DataFormatter();
								System.out.println(formatter.
										formatRawCellContents(
												cellOutput.getNumericCellValue(), 
												style.getDataFormat(),
												style.getDataFormatString())
										);
							}
							break;
						case HSSFCell.CELL_TYPE_BOOLEAN:
							System.out.println(cellOutput.getBooleanCellValue());
							break;
						case HSSFCell.CELL_TYPE_ERROR:
							System.out.println(ErrorEval.getText(cellOutput.getErrorCellValue()));							
							break;
					}
										
					break;
			}							
		}			
						
		workbook.close();		
	}
}