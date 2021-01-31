package excel.reading;

import java.io.File;
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

public class ExcelReading {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		readExcelFormula();
	}

	public static void readExcelFormula() {
		ObjectMapper mapper = new ObjectMapper();

		// assuming xlsx file
		Workbook workbook = null;
		try {
			File file = new File("C:\\Users\\SESA452171\\Documents\\ConvertJson.xlsx");

			OPCPackage opcPackage = OPCPackage.open(file);
			workbook = new XSSFWorkbook(opcPackage);
			Sheet sheet1 = workbook.getSheet("code_description");
			Sheet sheet2 = workbook.getSheet("code_details");
			int counter = 0;
			int counter1 = 0;

			List<SelectedKeyJSONObject> listSelectedKeyJSONObject = new ArrayList<SelectedKeyJSONObject>();
			for (Row row1 : sheet1) {
				if (counter > 0) {

					SelectedKeyJSONObject selectedKeyJSONObject = new SelectedKeyJSONObject();

					Cell cell1 = row1.getCell(0);
					Cell cell2 = row1.getCell(1);
					selectedKeyJSONObject.selectionKeyName = getCellValue(cell2).toString();

					
					selectedKeyJSONObject.selectionCategory = getCellValue(cell1).toString();

					Cell cell3 = row1.getCell(2);
					selectedKeyJSONObject.selectionKeyDescription = getCellValue(cell3).toString();

					listSelectedKeyJSONObject.add(selectedKeyJSONObject);

				}
				counter++;
			}

			for (Row row2 : sheet2) {

				for (int k = 0; k < listSelectedKeyJSONObject.size(); k++) {
					if (counter1 > 0) {

						Cell cell1 = row2.getCell(0);
						String selectionKeyName = getCellValue(cell1).toString();

						Cell cell2 = row2.getCell(1);
						String subAssemblyID = getCellValue(cell2).toString();

						Cell cell3 = row2.getCell(2);
						String subAssemblyQuantity = getCellValue(cell3).toString();

						SubAssemblyDetail subAssembly = new SubAssemblyDetail();
						subAssembly.setSubAssemblyID(subAssemblyID);

						if (!subAssemblyQuantity.equals("CTD")) {
							subAssembly.setSubAssemblyQuantity(Float.parseFloat(subAssemblyQuantity));
						}
						if (selectionKeyName.equals(listSelectedKeyJSONObject.get(k).selectionKeyName)) {

							if (listSelectedKeyJSONObject.get(k).subAssemblyIDList == null) {
								ArrayList<SubAssemblyDetail> SLIst = new ArrayList<SubAssemblyDetail>();
								SLIst.add(subAssembly);
								listSelectedKeyJSONObject.get(k).subAssemblyIDList = SLIst;
							} else {
								listSelectedKeyJSONObject.get(k).subAssemblyIDList.add(subAssembly);
							}

						}
					}
					counter1++;
				}

			}

			StringBuilder writeUp = new StringBuilder();

			writeUp.append("[");

			for (int i = 0; i < listSelectedKeyJSONObject.size(); i++) {

				StringBuilder newWriteup = new StringBuilder();
				newWriteup.append("{ \"selectionKeyName\" :\"")
						.append(listSelectedKeyJSONObject.get(i).selectionKeyName).append("\",")
						.append("\"selectionKeyDescription\" : \"")
						.append(listSelectedKeyJSONObject.get(i).selectionKeyDescription).append("\",")
						.append("\"selectionCategory\" : \"").append(listSelectedKeyJSONObject.get(i).selectionCategory)
						.append("\",").append("\"subAssemblyIDList\" : ");
				StringBuilder subList = new StringBuilder();
				System.out.print(listSelectedKeyJSONObject.get(i).selectionKeyName);
				if (listSelectedKeyJSONObject.get(i).subAssemblyIDList != null) {
					for (int l = 0; l < listSelectedKeyJSONObject.get(i).subAssemblyIDList.size(); l++) {
						System.out.print("i" + i + "\n");
						System.out
								.print("l" + l + listSelectedKeyJSONObject.get(i).subAssemblyIDList.toString() + "\n");
						subList.append("{\"subAssemblyID\" :\"")
								.append(listSelectedKeyJSONObject.get(i).subAssemblyIDList.get(l).getSubAssemblyID())
								.append("\",").append("\"subAssemblyQuantity\":")
								.append(listSelectedKeyJSONObject.get(i).subAssemblyIDList.get(l)
										.getSubAssemblyQuantity())
								.append("},");
						
					}
				}

				newWriteup.append("[").append(subList).append("]},");
				writeUp.append(newWriteup);
			}

			writeUp.append("]");

			mapper.writerWithDefaultPrettyPrinter()
					.writeValue(new File("C:\\Users\\SESA452171\\Desktop\\masterData.txt"), writeUp.toString());

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
