package excel.reading;

import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.fasterxml.jackson.core.JsonGenerationException;
import com.fasterxml.jackson.databind.JsonMappingException;
import com.fasterxml.jackson.databind.ObjectMapper;

import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.util.Iterator;
  
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
import org.json.simple.parser.ParseException;

public class ConstraintsToJson {
	
	private static JSONArray currentJSONKeys;

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		getConstraintsKey();
		readConstraintsExcel();
		
	}

	@SuppressWarnings("unchecked")
	public static void getConstraintsKey(){		
		try {
			JSONParser parser = new JSONParser();
			Object obj = parser.parse(new FileReader("C:\\Users\\SESA547061\\Desktop\\constraintsKey.json"));			
			currentJSONKeys =  (JSONArray) obj;	
			System.out.println(currentJSONKeys);				
			
		} catch (Exception ex) {
            ex.printStackTrace();
        }

	}

	@SuppressWarnings("unlikely-arg-type")
	public static void readConstraintsExcel() {
		ObjectMapper mapper = new ObjectMapper();

		// assuming xlsx file
		Workbook workbook = null;
		List<RackDetails> rackDetailsList = new ArrayList<RackDetails>();
		try {
			File file = new File("C:\\Users\\SESA547061\\Desktop\\ConstraintsToJSON.xlsx");

			OPCPackage opcPackage = OPCPackage.open(file);
			workbook = new XSSFWorkbook(opcPackage);
			/*
			 * Rack Details Constaraints to json
			 */

			Sheet sheet1 = workbook.getSheet("Rack");

			int counterRackDetails = 0;
			for (Row row1 : sheet1) {
				if (counterRackDetails > 0) {

					RackDetails rackDetails = new RackDetails();
					ValueObject value = new ValueObject();

					Cell cell0 = row1.getCell(0);
					Cell cell1 = row1.getCell(1);
					Cell cell2 = row1.getCell(2);
					Cell cell3 = row1.getCell(3);
					Cell cell4 = row1.getCell(4);
					Cell cell5 = row1.getCell(5);
					Cell cell6 = row1.getCell(6);
					Cell cell7 = row1.getCell(7);
					rackDetails.rackId = getCellValue(cell0).toString();
					rackDetails.height = Double.valueOf(getCellValue(cell1).toString());
					rackDetails.width = Double.valueOf(getCellValue(cell2).toString());
					rackDetails.depth = Double.valueOf(getCellValue(cell3).toString());
					value.Inrow_Container = Double.valueOf(getCellValue(cell4).toString());
					value.Inrow_Module = Double.valueOf(getCellValue(cell5).toString());
					value.Overhead_Container = Double.valueOf(getCellValue(cell6).toString());
					value.Overhead_Module = Double.valueOf(getCellValue(cell7).toString());
					rackDetails.value = value;

					rackDetailsList.add(rackDetails);

				}
				counterRackDetails++;
			}

			StringBuilder writeUpRack = new StringBuilder();

			writeUpRack.append("[");

			for (int i = 0; i < rackDetailsList.size(); i++) {

				writeUpRack.append("{ \"rackId\" :\"").append(rackDetailsList.get(i).rackId).append("\"").append(",")
						.append("\"height\" : ").append(rackDetailsList.get(i).height).append(",")
						.append("\"width\" : ").append(rackDetailsList.get(i).width).append(",")
						.append("\"depth\" : ").append(rackDetailsList.get(i).depth).append(",")
						.append("\"value\" : ").append("{").append("\"Inrow_Container\" : ")
						.append(rackDetailsList.get(i).value.Inrow_Container).append(",")
						.append("\"Overhead_Container\" : ").append(rackDetailsList.get(i).value.Overhead_Container)
						.append(",").append("\"Inrow_Module\" : ").append(rackDetailsList.get(i).value.Inrow_Module)
						.append(",").append("\"Overhead_Module\" : ")
						.append(rackDetailsList.get(i).value.Overhead_Module);
				writeUpRack.append("}},");
			}

			writeUpRack.append("]");

			mapper.writerWithDefaultPrettyPrinter()
					.writeValue(new File("C:\\Users\\SESA547061\\Desktop\\rackDetailsconstraintsJson.txt"), writeUpRack.toString());
			/*
			 * rackDetails constraints creation ends here 
			 */
			
			
			
			/*
			 * pduconstraintsJson Constaraints to json
			 */

			Sheet sheet2 = workbook.getSheet("PDU");
			List<ModelDetiails> pduModelkDetailsList = new ArrayList<ModelDetiails>();

			int counterPDU = 0;
			for (Row row1 : sheet2) {
				if (counterPDU > 0) {

					ModelDetiails pduModelkDetails = new ModelDetiails();

					Cell cell0 = row1.getCell(0);
					Cell cell1 = row1.getCell(1);
					Cell cell2 = row1.getCell(2);
					pduModelkDetails.modelType = getCellValue(cell0).toString();
					pduModelkDetails.modelNumber = getCellValue(cell1).toString();
					pduModelkDetails.modelDescription = getCellValue(cell2).toString();
					pduModelkDetailsList.add(pduModelkDetails);

				}
				counterPDU++;
			}

			StringBuilder writeUp = new StringBuilder();

			writeUp.append("{");

			for (int i = 0; i < pduModelkDetailsList.size(); i++) {

				writeUp.append("\"").append(pduModelkDetailsList.get(i).modelType).append("\" : {").
				append(" \"model\" :\"").append(pduModelkDetailsList.get(i).modelNumber).append("\"").append(",")
						.append("\"description\" : \"").append(pduModelkDetailsList.get(i).getModelDescription()).append("\"");

				writeUp.append("},");
			}

			writeUp.append("}");

			mapper.writerWithDefaultPrettyPrinter()
					.writeValue(new File("C:\\Users\\SESA547061\\Desktop\\pduconstraintsJson.txt"), writeUp.toString());
			/*
			 * pduconstraintsJson constraints creation ends here 
			 */
			
			
			/*
			 * noupsconstraintsJson Constaraints to json
			 */
            for(int sheetNum = 0; sheetNum < workbook.getNumberOfSheets(); sheetNum++) {
            	if(workbook.getSheetName(sheetNum).toString().toLowerCase().equals("NoUPS".toLowerCase())) {
					Sheet sheet3 = workbook.getSheetAt(sheetNum);						
					List<Map<Integer, String>> modelDetailsList = new ArrayList<Map<Integer, String>>();			
					
					for(int rowNumber = 0; rowNumber < sheet3.getLastRowNum(); rowNumber++) {
					    Row row = sheet3.getRow(rowNumber);
					    Map<Integer, String> modelDetails = new HashMap<Integer, String>();			    
					    for(int columnNumber = 0; columnNumber < row.getLastCellNum(); columnNumber++) {
					        Cell cell = row.getCell(columnNumber);			        
					        if(cell == null || getCellValue(cell) == null) {
					        	modelDetails.put(columnNumber, Integer.toString(0));
					        }
					        else {
					        	modelDetails.put(columnNumber, getCellValue(cell).toString());
					        }			        
					    }
					    modelDetailsList.add(modelDetails);
					}
		
					StringBuilder writeUpData = new StringBuilder();			
					writeUpData.append("[");
					for (int i = 1; i < modelDetailsList.size(); i++) {		
						Map<Integer, String> modelDetails = modelDetailsList.get(i);
									
						writeUpData.append("{");
						for(int j=0; j < modelDetails.size(); j++) {	
							writeUpData.append("\"");
							writeUpData.append(modelDetailsList.get(0).get(j));
							writeUpData.append("\" :\"");
							writeUpData.append(modelDetailsList.get(i).get(j));
							writeUpData.append("\",");					
						}
						
						writeUpData.append("},");				
					}			
		
					writeUpData.append("]");
					
					String targetFile = "C:\\\\Users\\\\SESA547061\\\\Desktop\\\\" + workbook.getSheetName(sheetNum) + "ConstraintsJson.txt";
		
					mapper.writerWithDefaultPrettyPrinter()
							.writeValue(new File(targetFile), writeUpData.toString());
					
		            }
            }
			/*
			 * noupsconstraintsJson constraints creation ends here 
			 */
			
			
			
		} catch (JsonGenerationException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (JsonMappingException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException | InvalidFormatException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		System.out.print(rackDetailsList);

	}
	

	
	
	
	

	private static Object getCellValue(Cell cell) {
		Object cellValue = null;
		switch (cell.getCellType().toString()) {
		case "BOOLEAN":
			cellValue = cell.getBooleanCellValue();
			break;
		case "NUMERIC":
			if (DateUtil.isCellDateFormatted(cell)) {
				cellValue = cell.getDateCellValue();
			} else {
				cellValue = cell.getNumericCellValue();
			}
			break;
		case "STRING":
			cellValue = cell.getStringCellValue();
			System.out.println(cellValue);
			break;
		case "FORMULA":
			System.out.println("Cell Formula=" + cell.getCellFormula());
			System.out.println("Cell Formula Result Type=" + cell.getCachedFormulaResultType());

			if (cell.getCachedFormulaResultType().toString() == "NUMERIC") {
				System.out.println("Formula Value=" + cell.getNumericCellValue());
				cellValue = cell.getNumericCellValue();
			}
			if (cell.getCachedFormulaResultType().toString() == "STRING") {
				System.out.println("Formula Value=" + cell.getStringCellValue());
				cellValue = cell.getStringCellValue();
			}
			break;
		case "BLANK":
			cellValue = "";
			break;
		default:
			cellValue = "";
		}
		return cellValue;
	}

}
