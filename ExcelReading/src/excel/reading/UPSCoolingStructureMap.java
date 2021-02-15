package excel.reading;

import java.util.List;

public class UPSCoolingStructureMap {
	
	String UPSFamilyRedundancy ;
	List<CoolingStructureMap> coolingStructureMapList;
	
	
	public String getUPSFamilyRedundancy() {
		return UPSFamilyRedundancy;
	}
	public void setUPSFamilyRedundancy(String uPSFamilyRedundancy) {
		UPSFamilyRedundancy = uPSFamilyRedundancy;
	}
	public List<CoolingStructureMap> getCoolingStructureMapList() {
		return coolingStructureMapList;
	}
	public void setCoolingStructureMapList(List<CoolingStructureMap> coolingStructureMapList) {
		this.coolingStructureMapList = coolingStructureMapList;
	}
	

}
