package excel.reading;

import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;

import com.fasterxml.jackson.core.JsonGenerationException;
import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.JsonMappingException;
import com.fasterxml.jackson.databind.ObjectMapper;

public class ConstraintsToJson {

	private static JSONObject currentJSONKeys;
	private static String filePath = new File("").getAbsolutePath();

	public static void main(String[] args) {

		ConstraintsToJson ConstraintsToJson = new ConstraintsToJson();
		// TODO Auto-generated method stub
		ConstraintsToJson.getConstraintsKey();
		ConstraintsToJson.readConstraintsExcel();
		


	}

	@SuppressWarnings("unchecked")
	public void getConstraintsKey() {
		try {
			JSONParser parser = new JSONParser();
			currentJSONKeys = (JSONObject) parser
					.parse(new FileReader(filePath + "\\inputResources\\constraintsKey.json"));
			System.out.println(currentJSONKeys);
		} catch (Exception ex) {
			ex.printStackTrace();
		}

	}

	public void readConstraintsExcel() {
		ObjectMapper mapper = new ObjectMapper();
		// assuming xlsx file
		Workbook workbook = null;

		try {

			File file = new File(filePath + "\\inputResources\\ConstraintsToJSON.xlsx");
			OPCPackage opcPackage = OPCPackage.open(file);
			workbook = new XSSFWorkbook(opcPackage);

			for (int sheetNum = 0; sheetNum < workbook.getNumberOfSheets(); sheetNum++) {
				String writeUpData = null;
				String targetFile = filePath + "\\outputResources\\" + workbook.getSheetName(sheetNum)
						+ "ConstraintsJson.txt";
				switch (workbook.getSheetName(sheetNum)) {
				case "NoUPS":
					writeUpData = readNoUPSConstraintsExcel(workbook, sheetNum);
					mapper.writerWithDefaultPrettyPrinter().writeValue(new File(targetFile), writeUpData.toString());
					break;
				case "Rack":
					writeUpData = readRackConstraintsExcel(workbook, sheetNum);
					mapper.writerWithDefaultPrettyPrinter().writeValue(new File(targetFile), writeUpData.toString());
					break;
				case "RackDetails":
					writeUpData = readRackDetasilConstraintsExcel(workbook, sheetNum);
					mapper.writerWithDefaultPrettyPrinter().writeValue(new File(targetFile), writeUpData.toString());
					break;
				case "PDU":
					writeUpData = readPDUConstraintsExcel(workbook, sheetNum);
					mapper.writerWithDefaultPrettyPrinter().writeValue(new File(targetFile), writeUpData.toString());
					break;
				case "UPS":
					writeUpData = readUPSDetailsConstraintsExcel(workbook, sheetNum);
					mapper.writerWithDefaultPrettyPrinter().writeValue(new File(targetFile), writeUpData.toString());
					break;
				case "ContainerUPS":
					writeUpData = readContainerUPSConstraintsExcel(workbook, sheetNum);
					mapper.writerWithDefaultPrettyPrinter().writeValue(new File(targetFile), writeUpData.toString());
					break;

				default:
					break;
				}
			}

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

	}

	private String readContainerUPSConstraintsExcel(Workbook workbook, int sheetNum) {
		// TODO Auto-generated method stub
		Sheet sheet = workbook.getSheetAt(sheetNum);
		Map<String, List<CoolingStructureMap>> uPSCoolingStructureMap = new HashMap<String, List<CoolingStructureMap>>();
		List<String> coolingTypesList = new ArrayList<String>();
		for (Row row : sheet) {

			if (row.getCell(0) != null && !getCellValue(row.getCell(0)).toString().equals("")
					&& (!getCellValue(row.getCell(0)).toString().contains("ISO"))) {
				String coolingType = getCellValue(row.getCell(1)).toString();
				String coolingID = getCellValue(row.getCell(0)).toString();
				Double coolinglength = 0.0;
				String CoolingTypeValue = "";

				if (coolingType.contains("INROW") && coolingType.contains("CW")) {
					coolinglength = Double.valueOf(getCellValue(row.getCell(3)).toString());
					CoolingTypeValue = "Inrow CW " + coolinglength + "mm-" + coolingID;
				} else if (coolingType.contains("INROW") && coolingType.contains("DX")) {
					coolinglength = Double.valueOf(getCellValue(row.getCell(3)).toString());
					CoolingTypeValue = "Inrow DX " + coolinglength + "mm-" + coolingID;
				} else if (coolingType.contains("CRAC") && coolingType.contains("DX")) {
					CoolingTypeValue = "CRAC DX " + coolingID;
				} else if (coolingType.contains("CRAC") && coolingType.contains("CW")) {
					CoolingTypeValue = "CRAC CW " + coolingID;
				} else if (coolingType.contains("CRAH") && coolingType.contains("CW")) {
					CoolingTypeValue = "CRAH CW-" + coolingID;
				} else if (coolingType.contains("CRAH") && coolingType.contains("DX")) {
					CoolingTypeValue = "CRAH DX-" + coolingID;
				} else if (coolingType.contains("WALLMOUNT")) {
					CoolingTypeValue = "Wall Mounted Down Flow" + coolingID;
				} else if (coolingType.contains("UNISPLIT")) {
					CoolingTypeValue = "Unisplit DX " + coolingID;
				}

				if (!CoolingTypeValue.equals("")) {
					coolingTypesList.add(new String(CoolingTypeValue));
				}

			}

			else if (row.getCell(0) != null && !getCellValue(row.getCell(0)).toString().equals("")
					&& (getCellValue(row.getCell(0)).toString().contains("ISO"))) {
				String LayoutRedundancy = getCellValue(row.getCell(1)).toString() + " "
						+ getCellValue(row.getCell(2)).toString();
				// UPSCoolingStructureMap uPSCoolingStructureMap = new UPSCoolingStructureMap();
				List<CoolingStructureMap> CoolingStructureMapList = new ArrayList<CoolingStructureMap>();

				for (int j = 0; j < coolingTypesList.size(); j++) {
					CoolingStructureMap coolingStructureMap = new CoolingStructureMap();
					String coolType = coolingTypesList.get(j);

					coolingStructureMap.coolingType = coolingTypesList.get(j);
					List<StructureDetails> StructureDetailsList = new ArrayList<StructureDetails>();
					coolingStructureMap.setStructureDetailsList(StructureDetailsList);
					CoolingStructureMapList.add(coolingStructureMap);
				}

				if (uPSCoolingStructureMap.get(LayoutRedundancy) == null) {
					uPSCoolingStructureMap.put(LayoutRedundancy, CoolingStructureMapList);
				}

				List<CoolingStructureMap> coolingStructureMapListActual = uPSCoolingStructureMap.get(LayoutRedundancy);
				int cellNumber = 4;

				// System.out.print("size" + size);

				for (int j = 0; j < coolingStructureMapListActual.size(); j++) {
					CoolingStructureMap CoolingStructureMap = coolingStructureMapListActual.get(j);

					List<StructureDetails> StructureDetailsList = CoolingStructureMap.getStructureDetailsList();
					StructureDetails structureDetails = new StructureDetails();

					// Common hardcoded values go here..
					structureDetails.setStructureValue(0);
					structureDetails.setItLoad(0);
					structureDetails.setMinimumServiceLength(0);
					structureDetails.setElectricalPanel(0);
				
					

					if (StructureDetailsList.size() >= 4) // means dual bay
					{
						structureDetails.setLength(Double.valueOf(getCellValue(row.getCell(cellNumber)).toString()));
						structureDetails.setType(getCellValue(row.getCell(0)).toString());
						structureDetails.setBayType("Dual");
						if(!CoolingStructureMap.coolingType.contains("CRA"))
						{
							int newCell1 =  cellNumber + (coolingTypesList.size() - cellNumber)  + (coolingTypesList.size() - 2) + 1 ;
							structureDetails.setValue(row.getCell(newCell1) == null  ? 0.0 :  Double.valueOf(getCellValue(row.getCell(cellNumber + coolingTypesList.size())).toString()));
							
						}
						else
						{
						structureDetails.setValue(0);	
						int newCell1 =  cellNumber + (coolingTypesList.size() + 4 - cellNumber)  + (coolingTypesList.size() - 2) + 2 ;
						int newCell2 =  cellNumber + (coolingTypesList.size() + 4 - cellNumber)  + (coolingTypesList.size() - 2) + 4;
						structureDetails.setMinimumServiceLength(Double.valueOf(getCellValue(row.getCell(newCell1)).toString()) );
						structureDetails.setElectricalPanel(Double.valueOf(getCellValue(row.getCell(newCell2)).toString()) );
						}
						cellNumber++;

					} else {
						if (!CoolingStructureMap.coolingType.contains("CRA")) {
							structureDetails
									.setLength(Double.valueOf(getCellValue(row.getCell(cellNumber)).toString()));
							structureDetails.setType(getCellValue(row.getCell(0)).toString());
							structureDetails.setBayType("Single");
							int newCell = cellNumber + coolingTypesList.size();
							structureDetails.setValue(row.getCell(newCell) == null  ? 0.0 :  Double.valueOf(getCellValue(row.getCell(cellNumber + coolingTypesList.size())).toString()));
							cellNumber++;
						} else {

							structureDetails.setBayType("Single");
							structureDetails.setLength(0);
							structureDetails.setType(getCellValue(row.getCell(0)).toString());
							structureDetails.setValue(0);
						}
					}
					if (structureDetails.getType().contains("NON")) {
						structureDetails.setStructureType("Module");
					} else {
						structureDetails.setStructureType("Container");
					}
					StructureDetailsList.add(structureDetails);
					CoolingStructureMap.setStructureDetailsList(StructureDetailsList);

				}
			}

		}
		System.out.println("final" + uPSCoolingStructureMap);
		StringBuilder writeUpRack = new StringBuilder();

		writeUpRack.append("[");
		
		
		Iterator<Map.Entry<String, List<CoolingStructureMap>>> iterator = uPSCoolingStructureMap.entrySet().iterator();
		while (iterator.hasNext()) {
			Map.Entry<String, List<CoolingStructureMap>> entry = iterator.next();
			System.out.println(entry.getKey() + ":" + entry.getValue());

			writeUpRack.append("{ \"upsFamilyRedundancy\" :\"").append(entry.getKey()).append("\"").append(",")
					.append("\"cooling\" : [");

			List<CoolingStructureMap> coolonMapList = entry.getValue();
			for (int k = 0; k < coolonMapList.size(); k++) {
				CoolingStructureMap coolingStructureMap = coolonMapList.get(k);
				String coolingType = coolingStructureMap.getCoolingType();
				writeUpRack.append("{ \"coolingType\" :\"").append(coolingType).append("\",")
						.append("\"StructureDetails \" : [");

				for (int m = 0; m < coolingStructureMap.getStructureDetailsList().size(); m++)

				{
					writeUpRack.append("{");
					StructureDetails structureDetails = coolingStructureMap.getStructureDetailsList().get(m);
					writeUpRack.append(" \"type\" : \"").append(structureDetails.getType()).append("\" ,")
							.append(" \"length\" : ").append(structureDetails.getLength()).append(" ,")
							.append(" \"bayType\" : \"").append(structureDetails.getBayType()).append("\" ,")
							.append(" \"value\" : ").append(structureDetails.getValue()).append(" ,")
							.append(" \"structureType\" : \"").append(structureDetails.getStructureType())
							.append("\" ,").append(" \"structureValue\" : ")
							.append(structureDetails.getStructureValue()).append(" ,")
							.append(" \"dehumidifier\" : ").append(structureDetails.getDehumidifier()).append(" ,")
							.append(" \"minimumServiceLength\" : ").append(structureDetails.getMinimumServiceLength())
							.append(" ,").append(" \"electricalPanel\" : ")
							.append(structureDetails.getElectricalPanel()).append(" ,").append(" \"itLoad\" : ")
							.append(structureDetails.getItLoad()).append(" },");

				}
				writeUpRack.append("]},");

			}
			writeUpRack.append("]},");
		}

		writeUpRack.append("]");

		return writeUpRack.toString();

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

	public static String readNoUPSConstraintsExcel(Workbook workbook, int sheetNum) {
		Sheet sheet = workbook.getSheetAt(sheetNum);
		ArrayList<HashMap<String, String>> modelDetailsList = new ArrayList<HashMap<String, String>>();
		JSONObject mappingKeys = (JSONObject) currentJSONKeys.get(workbook.getSheetName(sheetNum));
		System.out.println(mappingKeys);

		for (int rowNumber = 1; rowNumber < sheet.getLastRowNum(); rowNumber++) {
			Row row = sheet.getRow(rowNumber);
			HashMap<String, String> modelDetails = new HashMap<String, String>();
			for (int columnNumber = 0; columnNumber < row.getLastCellNum(); columnNumber++) {
				Cell cell = row.getCell(columnNumber);
				if (cell == null || getCellValue(cell) == null || getCellValue(cell).toString().isEmpty()) {
					modelDetails.put(mappingKeys.get(Integer.toString(columnNumber)).toString(), Integer.toString(0));
				} else {
					modelDetails.put(mappingKeys.get(Integer.toString(columnNumber)).toString(),
							getCellValue(cell).toString());
				}
			}
			modelDetailsList.add(modelDetails);
		}

		ObjectMapper objectMapper = new ObjectMapper();
		String finalJsonString;
		try {
			finalJsonString = objectMapper.writeValueAsString(modelDetailsList);
			System.out.println(finalJsonString);
			return finalJsonString;
		} catch (JsonProcessingException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return "";
	}

	public static String readRackConstraintsExcel(Workbook workbook, int sheetNum) {
		Sheet sheet = workbook.getSheetAt(sheetNum);
		List<Map<Integer, String>> modelDetailsList = new ArrayList<Map<Integer, String>>();
		Object mappingKeys = currentJSONKeys.get(workbook.getSheetName(sheetNum));
		System.out.println(mappingKeys);

		for (int rowNumber = 0; rowNumber < sheet.getLastRowNum(); rowNumber++) {
			Row row = sheet.getRow(rowNumber);
			Map<Integer, String> modelDetails = new HashMap<Integer, String>();
			for (int columnNumber = 0; columnNumber < row.getLastCellNum(); columnNumber++) {
				Cell cell = row.getCell(columnNumber);
				if (cell == null || getCellValue(cell) == null) {
					modelDetails.put(columnNumber, Integer.toString(0));
				} else {
					modelDetails.put(columnNumber, getCellValue(cell).toString());
				}
			}
			modelDetailsList.add(modelDetails);
		}

		StringBuilder writeUpData = new StringBuilder();
		writeUpData.append("[{");
		for (int i = 1; i < modelDetailsList.size(); i++) {
			writeUpData.append("{ \"");
			writeUpData.append(modelDetailsList.get(i).get(0));
			writeUpData.append("\": {");
			writeUpData.append("\"model\" :\"");
			writeUpData
					.append(modelDetailsList.get(i).get(0).substring(modelDetailsList.get(i).get(0).indexOf("_") + 1));
			writeUpData.append("\", \"");
			writeUpData.append("\"description\" :\"");
			writeUpData.append(modelDetailsList.get(i).get(1));
			writeUpData.append("\"},");
		}
		writeUpData.append("}");
		writeUpData.append("]");
		return writeUpData.toString();
	}

	public static String readRackDetasilConstraintsExcel(Workbook workbook, int sheetNum) {
		Sheet sheet = workbook.getSheet("RackDetails");
		List<RackDetails> rackDetailsList = new ArrayList<RackDetails>();
		int counterRackDetails = 0;
		for (Row row1 : sheet) {
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
					.append("\"height\" : ").append(rackDetailsList.get(i).height).append(",").append("\"width\" : ")
					.append(rackDetailsList.get(i).width).append(",").append("\"depth\" : ")
					.append(rackDetailsList.get(i).depth).append(",").append("\"value\" : ").append("{")
					.append("\"Inrow_Container\" : ").append(rackDetailsList.get(i).value.Inrow_Container).append(",")
					.append("\"Overhead_Container\" : ").append(rackDetailsList.get(i).value.Overhead_Container)
					.append(",").append("\"Inrow_Module\" : ").append(rackDetailsList.get(i).value.Inrow_Module)
					.append(",").append("\"Overhead_Module\" : ").append(rackDetailsList.get(i).value.Overhead_Module);
			writeUpRack.append("}},");
		}

		writeUpRack.append("]");

		return writeUpRack.toString();
	}

	public static String readPDUConstraintsExcel(Workbook workbook, int sheetNum) {
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

			writeUp.append("\"").append(pduModelkDetailsList.get(i).modelType).append("\" : {").append(" \"model\" :\"")
					.append(pduModelkDetailsList.get(i).modelNumber).append("\"").append(",")
					.append("\"description\" : \"").append(pduModelkDetailsList.get(i).getModelDescription())
					.append("\"");

			writeUp.append("},");
		}
		writeUp.append("}");

		return writeUp.toString();

	}

	public static String readUPSDetailsConstraintsExcel(Workbook workbook, int sheetNum) {
		Sheet sheet = workbook.getSheetAt(sheetNum);
		ArrayList<HashMap<String, String>> modelDetailsList = new ArrayList<HashMap<String, String>>();
		JSONObject mappingKeys = (JSONObject) currentJSONKeys.get(workbook.getSheetName(sheetNum));
		System.out.println(mappingKeys);
		boolean IsResetRow = true;
		for (int rowNumber = 0; rowNumber < sheet.getLastRowNum(); rowNumber++) {
			Row row = sheet.getRow(rowNumber);
			if (row.getCell(0).toString().equals("GALAXY VS/VM")) {
				break;
			}
			if (row.getCell(0).toString().equals("SYMMETRA")) {
				IsResetRow = true;
				continue;
			}

			for (int columnNumber = 0; columnNumber < row.getLastCellNum(); columnNumber++) {
				Cell cell = row.getCell(columnNumber);
				if (IsResetRow) {
					HashMap<String, String> modelDetails = new HashMap<String, String>();
					if (cell == null || getCellValue(cell) == null || getCellValue(cell).toString().isEmpty()) {
						modelDetails.put(Integer.toString(rowNumber), Integer.toString(0));
					} else {
						modelDetails.put(Integer.toString(rowNumber), getCellValue(cell).toString());
					}
					modelDetailsList.add(modelDetails);
				} else {
					if (cell == null || getCellValue(cell) == null || getCellValue(cell).toString().isEmpty()) {
						modelDetailsList.get(columnNumber).put(Integer.toString(rowNumber), Integer.toString(0));
					} else {
						modelDetailsList.get(columnNumber).put(Integer.toString(rowNumber),
								getCellValue(cell).toString());
					}
				}
			}
			IsResetRow = false;

		}

		ObjectMapper objectMapper = new ObjectMapper();
		String finalJsonString;
		try {
			finalJsonString = objectMapper.writeValueAsString(modelDetailsList);
			System.out.println(finalJsonString);
			return finalJsonString;
		} catch (JsonProcessingException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return "";
	}

}
