import java.time.LocalDate;
import java.time.Year;
import java.util.List;

public class DynamicVariables {
    private String consultantCompanyLogo;
    private String endClientLogo;
    private String endClientName;
    private String reportId;
    private List<String> assetTypeList;
    private LocalDate createdOn;
    private String createBy;
    private Year year;
    private String sccName;
    private String sccAddress;
    private String sccEmailId;

    public String getConsultantCompanyLogo() {
        return consultantCompanyLogo;
    }

    public void setConsultantCompanyLogo(String consultantCompanyLogo) {
        this.consultantCompanyLogo = consultantCompanyLogo;
    }

    public String getEndClientLogo() {
        return endClientLogo;
    }

    public void setEndClientLogo(String endClientLogo) {
        this.endClientLogo = endClientLogo;
    }

    public String getEndClientName() {
        return endClientName;
    }

    public void setEndClientName(String endClientName) {
        this.endClientName = endClientName;
    }

    public String getReportId() {
        return reportId;
    }

    public void setReportId(String reportId) {
        this.reportId = reportId;
    }

    public List<String> getAssetTypeList() {
        return assetTypeList;
    }

    public void setAssetTypeList(List<String> assetTypeList) {
        this.assetTypeList = assetTypeList;
    }

    public LocalDate getCreatedOn() {
        return createdOn;
    }

    public void setCreatedOn(LocalDate createdOn) {
        this.createdOn = createdOn;
    }

    public String getCreateBy() {
        return createBy;
    }

    public void setCreateBy(String createBy) {
        this.createBy = createBy;
    }

    public String getYear() {
        return year.toString();
    }

    public void setYear(Year year) {
        this.year = year;
    }

    public String getSccName() {
        return sccName;
    }

    public void setSccName(String sccName) {
        this.sccName = sccName;
    }

    public String getSccAddress() {
        return sccAddress;
    }

    public void setSccAddress(String sccAddress) {
        this.sccAddress = sccAddress;
    }

    public String getSccEmailId() {
        return sccEmailId;
    }

    public void setSccEmailId(String sccEmailId) {
        this.sccEmailId = sccEmailId;
    }
}
