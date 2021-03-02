package excel.reading;

import java.util.List;

public class EnclosureCoolingStrcture {
	
	String containerType ;

	List<CoolingDetails> coolingDetailsList ;
	
	
	public String getContainerType() {
		return containerType;
	}
	public void setContainerType(String containerType) {
		this.containerType = containerType;
	}
	public List<CoolingDetails> getCoolingDetails() {
		return coolingDetailsList;
	}
	public void setCoolingDetails(List<CoolingDetails> coolingDetailsList) {
		this.coolingDetailsList = coolingDetailsList;
	}

}
