import java.util.List;

public class DataDTO {
    private DynamicVariables dynamicVariables;
    private List<VulnerabilityCountTable> vulnerabilityCountTableList;
    private List<CriticalityTable> criticalityTables;

    public DynamicVariables getDynamicVariables() {
        return dynamicVariables;
    }

    public void setDynamicVariables(DynamicVariables dynamicVariables) {
        this.dynamicVariables = dynamicVariables;
    }

    public List<VulnerabilityCountTable> getVulnerabilityCountTableList() {
        return vulnerabilityCountTableList;
    }

    public void setVulnerabilityCountTableList(List<VulnerabilityCountTable> vulnerabilityCountTableList) {
        this.vulnerabilityCountTableList = vulnerabilityCountTableList;
    }

    public List<CriticalityTable> getCriticalityTables() {
        return criticalityTables;
    }

    public void setCriticalityTables(List<CriticalityTable> criticalityTables) {
        this.criticalityTables = criticalityTables;
    }
}
