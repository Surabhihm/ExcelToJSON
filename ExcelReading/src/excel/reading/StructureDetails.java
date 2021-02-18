package excel.reading;

public class StructureDetails {
	
	String type;
	double length;
	String bayType;
	double value;
	double  structureValue;
	String structureType;
	double itLoad;
	public double getItLoad() {
		return itLoad;
	}
	public void setItLoad(double itLoad) {
		this.itLoad = itLoad;
	}
	public String getStructureType() {
		return structureType;
	}
	public void setStructureType(String structureType) {
		this.structureType = structureType;
	}
	double dehumidifier;
	double minimumServiceLength;
	double electricalPanel;
	
	public String getType() {
		return type;
	}
	public void setType(String type) {
		this.type = type;
	}
	public double getLength() {
		return length;
	}
	public void setLength(double length) {
		this.length = length;
	}
	public String getBayType() {
		return bayType;
	}
	public void setBayType(String bayType) {
		this.bayType = bayType;
	}
	public double getValue() {
		return value;
	}
	public void setValue(double value) {
		this.value = value;
	}
	public double getStructureValue() {
		return structureValue;
	}
	public void setStructureValue(double structureValue) {
		this.structureValue = structureValue;
	}
	public double getDehumidifier() {
		return dehumidifier;
	}
	public void setDehumidifier(double dehumidifier) {
		this.dehumidifier = dehumidifier;
	}
	public double getMinimumServiceLength() {
		return minimumServiceLength;
	}
	public void setMinimumServiceLength(double minimumServiceLength) {
		this.minimumServiceLength = minimumServiceLength;
	}
	public double getElectricalPanel() {
		return electricalPanel;
	}
	public void setElectricalPanel(double electricalPanel) {
		this.electricalPanel = electricalPanel;
	}


}
