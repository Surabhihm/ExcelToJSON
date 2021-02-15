package excel.reading;

import java.util.List;

public class CoolingStructureMap {
	
	String coolingType ;
	
	List<StructureDetails> structureDetailsList;
	
	
	public String getCoolingType() {
		return coolingType;
	}
	public void setCoolingType(String coolingType) {
		this.coolingType = coolingType;
	}
	public List<StructureDetails> getStructureDetailsList() {
		return structureDetailsList;
	}
	public void setStructureDetailsList(List<StructureDetails> structureDetailsList) {
		this.structureDetailsList = structureDetailsList;
	}
	

}
