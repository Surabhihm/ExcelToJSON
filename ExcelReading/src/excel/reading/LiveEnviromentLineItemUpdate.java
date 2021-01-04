package excel.reading;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.fasterxml.jackson.core.JsonGenerationException;
import com.fasterxml.jackson.databind.JsonMappingException;
import com.fasterxml.jackson.databind.ObjectMapper;

public class LiveEnviromentLineItemUpdate {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		readExcelFormula();
	}

	public static void readExcelFormula() {
		ObjectMapper mapper = new ObjectMapper();
		Map<String, UpdatesPlaceHolder> liveEnv_map = new HashMap<String, UpdatesPlaceHolder>();
		List<DeleteLineItem> deleteLineItem_list = new ArrayList<DeleteLineItem>();
		// assuming xlsx file
		Workbook workbook = null;
		try {
			File file = new File(
					"C:\\Users\\SESA390538\\Desktop\\LiveEnviromentUpdate Excels\\LiveEnviromentUpdate.xlsx");
			FileInputStream fis = new FileInputStream(file);
//			System.setProperty("org.apache.poi.util.POILogger", "org.apache.poi.util.SystemOutLogger");
//			  System.setProperty("poi.log.level", POILogger.INFO + "");
			workbook = new XSSFWorkbook(fis);
			Sheet lieveEnv_sheet = workbook.getSheet("Articles");
			Sheet sheet = workbook.getSheetAt(0);

			int rowCount = 1;
			for (Row row : lieveEnv_sheet) {
				if(rowCount > 1) {
					double replaceWithLineItemInternalId = 0.0;
					double lineItemInternalCode = row.getCell(1).getNumericCellValue();

					String usedInSubAssemblies[] = row.getCell(8).getStringCellValue().split(",");
					String oracleID = row.getCell(10).getStringCellValue();
//					if("ORPHAN LINEITEM".equals(usedInSubAssemblies[0])) {
//						/** NO Action is taken, The row is not valid to be processed; */
//					} else
						if ("".equals(oracleID)) {
						/** NO Action is taken, The row is not valid to be processed; */
					} else if ("TITLE_DELETE".equals(oracleID)) {
						DeleteLineItem deleteLineItem = new DeleteLineItem(lineItemInternalCode,oracleID, usedInSubAssemblies);
						deleteLineItem_list.add(deleteLineItem);
					} else {
						UpdatesPlaceHolder UpdatesPlaceHolder = null;
						if (liveEnv_map.containsKey(oracleID)) {
							UpdatesPlaceHolder = liveEnv_map.get(oracleID);
						} else {
							double replaceWithLineItemInternalId$ = 0.0;
							List<UpdateLineItem> updateLineItemList = new ArrayList<UpdateLineItem>();
							UpdatesPlaceHolder = new UpdatesPlaceHolder(replaceWithLineItemInternalId$, updateLineItemList);
							liveEnv_map.put(oracleID, UpdatesPlaceHolder);
						}

						UpdateLineItem updatelineItem = new UpdateLineItem(lineItemInternalCode, replaceWithLineItemInternalId,
								oracleID, usedInSubAssemblies);
						UpdatesPlaceHolder.getLineItems().add(updatelineItem);
					}
				}
				rowCount++;
			}

		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		// Find out which line item to replace with the exixting one .
//		In the group with same oracle ID ,  Line item which have internalCode hax max value should replace the others.
		Set<String> oracleId_set = liveEnv_map.keySet();
		Iterator<String> keyItr = oracleId_set.iterator();
		while (keyItr.hasNext()) {
			List<Double> lineItemInternalCodeList = new ArrayList<Double>();
			String key = keyItr.next();
			UpdatesPlaceHolder updatesPlaceHolder = liveEnv_map.get(key);
			List<UpdateLineItem> updateLineItemList = updatesPlaceHolder.getLineItems();
			Iterator<UpdateLineItem> updateLineItemItr = updateLineItemList.iterator();
			while (updateLineItemItr.hasNext()) {
				UpdateLineItem updatelineitem = updateLineItemItr.next();
				lineItemInternalCodeList.add(updatelineitem.getLineItemInternalCode());
			}

			Double replaceWithLineItemInternalId = Collections.max(lineItemInternalCodeList);
			updatesPlaceHolder.setReplaceWithLineItemInternalId(replaceWithLineItemInternalId);

			updateLineItemItr = updateLineItemList.iterator();
			while (updateLineItemItr.hasNext()) {
				UpdateLineItem updatelineitem = updateLineItemItr.next();
				updatelineitem.setReplaceWithLineItemInternalId(replaceWithLineItemInternalId);
				if (updatelineitem.getLineItemInternalCode() != replaceWithLineItemInternalId) {
					updatelineitem.setDeleteThisLineItem(true);
				}
			}
		}

		try {
			mapper.writerWithDefaultPrettyPrinter()
					.writeValue(new File(
							"C:\\Users\\SESA390538\\Desktop\\LiveEnviromentUpdate Excels\\LiveEnvUpdatedLineItems.txt"),
							liveEnv_map);
			mapper.writerWithDefaultPrettyPrinter()
			.writeValue(new File(
					"C:\\Users\\SESA390538\\Desktop\\LiveEnviromentUpdate Excels\\LiveEnvDeletedLineItems.txt"),
					deleteLineItem_list);
		} catch (JsonGenerationException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (JsonMappingException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}

class UpdatesPlaceHolder {

	private double replaceWithLineItemInternalId;

	private List<UpdateLineItem> lineItems;

	public UpdatesPlaceHolder(double replaceWithLineItemInternalId, List<UpdateLineItem> lineItems) {
		super();
		this.replaceWithLineItemInternalId = replaceWithLineItemInternalId;
		this.lineItems = lineItems;
	}

	public double getReplaceWithLineItemInternalId() {
		return replaceWithLineItemInternalId;
	}

	public void setReplaceWithLineItemInternalId(double replaceWithLineItemInternalId) {
		this.replaceWithLineItemInternalId = replaceWithLineItemInternalId;
	}

	public List<UpdateLineItem> getLineItems() {
		return lineItems;
	}

	public void setLineItems(List<UpdateLineItem> lineItems) {
		this.lineItems = lineItems;
	}
}

class UpdateLineItem {
	private double lineItemInternalCode;
	private double replaceWithLineItemInternalId;
	private String oracleID;
	private String[] usedInSubAssemblies;
	private boolean deleteThisLineItem = false;

	public UpdateLineItem(double lineItemInternalCode, double replaceWithLineItemInternalId, String oracleID,
			String[] usedInSubAssemblies) {
		super();
		this.lineItemInternalCode = lineItemInternalCode;
		this.replaceWithLineItemInternalId = replaceWithLineItemInternalId;
		this.oracleID = oracleID;
		this.usedInSubAssemblies = usedInSubAssemblies;
	}

	public double getLineItemInternalCode() {
		return lineItemInternalCode;
	}

	public void setLineItemInternalCode(double lineItemInternalCode) {
		this.lineItemInternalCode = lineItemInternalCode;
	}

	public double getReplaceWithLineItemInternalId() {
		return replaceWithLineItemInternalId;
	}

	public void setReplaceWithLineItemInternalId(double replaceWithLineItemInternalId) {
		this.replaceWithLineItemInternalId = replaceWithLineItemInternalId;
	}

	public String getOracleID() {
		return oracleID;
	}

	public void setOracleID(String oracleID) {
		this.oracleID = oracleID;
	}

	public String[] getUsedInSubAssemblies() {
		return usedInSubAssemblies;
	}

	public void setUsedInSubAssemblies(String[] usedInSubAssemblies) {
		this.usedInSubAssemblies = usedInSubAssemblies;
	}

	public boolean isDeleteThisLineItem() {
		return deleteThisLineItem;
	}

	public void setDeleteThisLineItem(boolean deleteThisLineItem) {
		this.deleteThisLineItem = deleteThisLineItem;
	}
}

class DeleteLineItem {
	private double lineItemInternalCode;
	private String oracleID;
	private String[] usedInSubAssemblies;

	public DeleteLineItem(double lineItemInternalCode, String oracleID, String[] usedInSubAssemblies) {
		super();
		this.lineItemInternalCode = lineItemInternalCode;
		this.oracleID = oracleID;
		this.usedInSubAssemblies = usedInSubAssemblies;
	}

	public double getLineItemInternalCode() {
		return lineItemInternalCode;
	}

	public void setLineItemInternalCode(double lineItemInternalCode) {
		this.lineItemInternalCode = lineItemInternalCode;
	}

	public String getOracleID() {
		return oracleID;
	}

	public void setOracleID(String oracleID) {
		this.oracleID = oracleID;
	}

	public String[] getUsedInSubAssemblies() {
		return usedInSubAssemblies;
	}

	public void setUsedInSubAssemblies(String[] usedInSubAssemblies) {
		this.usedInSubAssemblies = usedInSubAssemblies;
	}

}