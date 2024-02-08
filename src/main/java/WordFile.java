
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.Year;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class WordFile {

    private static LinkedHashMap<String, String> values = new LinkedHashMap<>();

    private static Pattern placeholderPattern = Pattern.compile("\\<[^<>]+\\>");
    public static void main(String[] args) throws IOException {
        WordFile wordFile = new WordFile();

        String inputFilePath = "C:\\Users\\qSEAp\\Documents\\Report\\VA-Report-pdf-Template.docx";
        String outputFilePath = "C:\\Users\\qSEAp\\Documents\\Report\\VA-WORD.docx";

        String clientLogoImagePath = "C:\\Users\\qSEAp\\Documents\\Report\\PDF\\Axis.png";
        String consultantLogoImagePath = "C:\\Users\\qSEAp\\Documents\\Report\\PDF\\QSeapLogo.png";
        DataDTO dataDTO = new DataDTO();
        DynamicVariables dynamicVariables = new DynamicVariables();
        dynamicVariables.setConsultantCompanyLogo("ConsultantLogo");
        dynamicVariables.setEndClientLogo("ClientLogo");
        dynamicVariables.setEndClientName("Axis Bank");
        dynamicVariables.setReportId("1801");
        dynamicVariables.setAssetTypeList(Arrays.asList("Asset1, Asset2, Asset3"));
        LocalDate date = LocalDate.of(2024, 01, 25);
        dynamicVariables.setCreatedOn(date);
        dynamicVariables.setCreateBy("Mourya");
        Year year = Year.of(2024);
        dynamicVariables.setYear(year);
        dynamicVariables.setSccName("qSEAp");
        dynamicVariables.setSccAddress("Navi Mumbai");
        dynamicVariables.setSccEmailId("qseap@qseap.com");

        dataDTO.setDynamicVariables(dynamicVariables);

        VulnerabilityCountTable row1 = new VulnerabilityCountTable();
        row1.setAssetTypeName("AssetType1");
        row1.setTotalVulnerabilites(12);

        VulnerabilityCountTable row2 = new VulnerabilityCountTable();
        row2.setAssetTypeName("AssetType2");
        row2.setTotalVulnerabilites(11);

        VulnerabilityCountTable row3 = new VulnerabilityCountTable();
        row3.setAssetTypeName("AssetType3");
        row3.setTotalVulnerabilites(10);

        VulnerabilityCountTable row4 = new VulnerabilityCountTable();
        row4.setAssetTypeName("AssetType4");
        row4.setTotalVulnerabilites(3);

        List<VulnerabilityCountTable> vulnerabilityCountTableList = new ArrayList<>();
        vulnerabilityCountTableList.add(row1);
        vulnerabilityCountTableList.add(row2);
        vulnerabilityCountTableList.add(row3);

        vulnerabilityCountTableList.add(row2);
        vulnerabilityCountTableList.add(row3);

        vulnerabilityCountTableList.add(row2);
        vulnerabilityCountTableList.add(row3);

        vulnerabilityCountTableList.add(row2);
        vulnerabilityCountTableList.add(row3);
        vulnerabilityCountTableList.add(row4);

        dataDTO.setVulnerabilityCountTableList(vulnerabilityCountTableList);

        int total = 0;

        CriticalityTable criticalityTableRow1 = new CriticalityTable();
        criticalityTableRow1.setGroupName(row1.getAssetTypeName());
        criticalityTableRow1.setCriticalCount(2);
        criticalityTableRow1.setHighCount(7);
        criticalityTableRow1.setMediumCount(2);
        criticalityTableRow1.setLowCount(3);
        criticalityTableRow1.setInfoCount(1);

        total = criticalityTableRow1.getCriticalCount() + criticalityTableRow1.getHighCount() + criticalityTableRow1.getMediumCount()
                + criticalityTableRow1.getLowCount() + criticalityTableRow1.getInfoCount();
        criticalityTableRow1.setTotal(total);


        CriticalityTable criticalityTableRow2 = new CriticalityTable();
        criticalityTableRow2.setGroupName(row2.getAssetTypeName());
        criticalityTableRow2.setCriticalCount(1);
        criticalityTableRow2.setHighCount(4);
        criticalityTableRow2.setMediumCount(3);
        criticalityTableRow2.setLowCount(3);
        criticalityTableRow2.setInfoCount(2);

        total = criticalityTableRow2.getCriticalCount() + criticalityTableRow2.getHighCount() + criticalityTableRow2.getMediumCount()
                + criticalityTableRow2.getLowCount() + criticalityTableRow2.getInfoCount();
        criticalityTableRow2.setTotal(total);

        CriticalityTable criticalityTableRow3 = new CriticalityTable();
        criticalityTableRow3.setGroupName(row3.getAssetTypeName());
        criticalityTableRow3.setCriticalCount(4);
        criticalityTableRow3.setHighCount(2);
        criticalityTableRow3.setMediumCount(1);
        criticalityTableRow3.setLowCount(7);
        criticalityTableRow3.setInfoCount(0);

        total = criticalityTableRow3.getCriticalCount() + criticalityTableRow3.getHighCount() + criticalityTableRow3.getMediumCount()
                + criticalityTableRow3.getLowCount() + criticalityTableRow3.getInfoCount();
        criticalityTableRow3.setTotal(total);


        List<CriticalityTable> criticalityTableList = new ArrayList<>();
        criticalityTableList.add(criticalityTableRow1);
        criticalityTableList.add(criticalityTableRow2);
        criticalityTableList.add(criticalityTableRow3);
        criticalityTableList.add(criticalityTableRow2);
        criticalityTableList.add(criticalityTableRow3);
        criticalityTableList.add(criticalityTableRow2);
        criticalityTableList.add(criticalityTableRow3);
        criticalityTableList.add(criticalityTableRow2);
        criticalityTableList.add(criticalityTableRow3);

        dataDTO.setCriticalityTables(criticalityTableList);

        List<CriticalityTable> staticRiskCountTable = new ArrayList<>();




        byte[] wordFilePath = convertDocumentToByteArray(inputFilePath);

        byte[] clientLogo = convertImageToByteArray(clientLogoImagePath);
        byte[] consultantLogo = convertImageToByteArray(consultantLogoImagePath);

//        values.put("<ConsultantCompanyLogo>", dynamicVariables.getConsultantCompanyLogo());
//        values.put("<EndClientLogo>", dynamicVariables.getEndClientLogo());
        values.put("<EndClientName>", dynamicVariables.getEndClientName());
        values.put("<ReportId>", dynamicVariables.getReportId());
        values.put("<AssetTypes>", String.join(", ", dynamicVariables.getAssetTypeList()));
        values.put("<CreatedOn>", String.valueOf(dynamicVariables.getCreatedOn()));
        values.put("<SecurityAuthorName>", dynamicVariables.getCreateBy());
        values.put("<Year>", String.valueOf(dynamicVariables.getYear()));
        values.put("<SCCName>", dynamicVariables.getSccName());
        values.put("<SCCAddress>", dynamicVariables.getSccAddress());
        values.put("<SCCEmail>", dynamicVariables.getSccEmailId());


        // Write data to Word file
        try {
            ByteArrayInputStream byteArrayInputStream = new ByteArrayInputStream(wordFilePath);
            XWPFDocument document = new XWPFDocument(byteArrayInputStream);


            wordFile.replacePlaceholders(document);
            wordFile.replaceDynamicValueWithImage(document, "<ConsultantCompanyLogo", consultantLogo);
            wordFile.replaceDynamicValueWithImage(document, "<EndClientLogo>", clientLogo);

            wordFile.populateVulnerabilityCountTable(document, vulnerabilityCountTableList);

            wordFile.populateCriticalityTable(document, criticalityTableList);



            //TODO:

            // Save the modified document to a new file
            FileOutputStream fos = new FileOutputStream(outputFilePath);
            document.write(fos);
            byteArrayInputStream.close();
            fos.close();
            System.out.println("Word file created successfully.");



        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

    }

    private void replacePlaceholders(XWPFDocument document) {
        for (XWPFParagraph paragraph : document.getParagraphs()) {
            replaceTextInParagraph(paragraph);
        }

        for (XWPFTable table : document.getTables()) {
            for (XWPFTableRow row : table.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph paragraph : cell.getParagraphs()) {
                        replaceTextInParagraph(paragraph);
                    }
                }
            }
        }

    }

    private void replaceTextInParagraph(XWPFParagraph paragraph) {
        List<XWPFRun> runs = paragraph.getRuns();
        for (int i = 0; i < runs.size(); i++) {
            XWPFRun run = runs.get(i);
            String text = run.getText(0);
            if (text != null) {
                Matcher matcher = placeholderPattern.matcher(text);

                while (matcher.find()) {
                    String placeholder = matcher.group(0);
                    System.out.println("Found placeholder: " + placeholder);
                    if (values.containsKey(placeholder)) {
                        text = text.replace(placeholder, values.get(placeholder));
                    }
                }
                run.setText(text, 0);
            }
        }
    }



    private void replaceDynamicValueWithImage(XWPFDocument document, String placeholder, byte[] imageBytes) {
        for (XWPFParagraph paragraph : document.getParagraphs()) {
            List<XWPFRun> runs = paragraph.getRuns();
            for (XWPFRun run : runs) {
                String text = run.getText(0);
                if (text != null && text.contains(placeholder)) {
                    // Clear existing run content
                    run.setText("", 0);

                    // Add image run from byte array
                    try (ByteArrayInputStream inputStream = new ByteArrayInputStream(imageBytes); ByteArrayInputStream logoStream = new ByteArrayInputStream(imageBytes)) {
                        run.addPicture(inputStream, XWPFDocument.PICTURE_TYPE_PNG, "Generated", Units.toEMU(75), Units.toEMU(25));
                    } catch (IOException | org.apache.poi.openxml4j.exceptions.InvalidFormatException e) {
                        e.printStackTrace();
                    }

                }
            }
        }
    }

    private void populateTableRow(XWPFTableRow row, VulnerabilityCountTable rowData) {
        List<XWPFTableCell> cells = row.getTableCells();
        if (cells.size() >= 2) { // Ensure there are at least two cells in the row
            // Populate the first cell with AssetTypeName
            XWPFTableCell assetTypeCell = cells.get(0);
            List<XWPFParagraph> assetTypeParagraphs = assetTypeCell.getParagraphs();
            if (assetTypeParagraphs.isEmpty()) {
                // Clear existing content and add new text
//                assetTypeCell.removeParagraph(0);
                XWPFParagraph assetTypeParagraph = assetTypeCell.addParagraph();
                assetTypeParagraph.createRun().setText(rowData.getAssetTypeName());
            } else {
                System.out.println("Error: The first cell does not contain any paragraphs.");
            }

            // Populate the second cell with TotalVul
            XWPFTableCell totalVulCell = cells.get(1);
            List<XWPFParagraph> totalVulParagraphs = totalVulCell.getParagraphs();
            if (totalVulParagraphs.isEmpty()) {
                // Clear existing content and add new text
//                totalVulCell.removeParagraph(0);
                XWPFParagraph totalVulParagraph = totalVulCell.addParagraph();
                totalVulParagraph.createRun().setText(String.valueOf(rowData.getTotalVulnerabilities()));
            } else {
                System.out.println("Error: The second cell does not contain any paragraphs.");
            }
        } else {
            System.out.println("Error: The table row does not contain the expected number of cells.");
        }
    }



    private void populateVulnerabilityCountTable(XWPFDocument document, List<VulnerabilityCountTable> vulnerabilityCountTableList) {
        // Find the target table (assuming it's the third table in the document)
        List<XWPFTable> tables = document.getTables();
        if (tables.size() >= 3) { // Ensure there are at least three tables in the document
            XWPFTable targetTable = tables.get(2); // Assuming the third table is the target table

            // Clear existing rows (excluding header row)
            int rowsToRemove = targetTable.getNumberOfRows() - 1;
            for (int i = 0; i < rowsToRemove; i++) {
                targetTable.removeRow(1);
            }

            // Add data to the table
            for (VulnerabilityCountTable rowData : vulnerabilityCountTableList) {
                XWPFTableRow row = targetTable.createRow();
                populateTableRow(row, rowData);
            }
        } else {
            System.out.println("Error: The document does not contain the expected number of tables.");
        }
    }

    private void populateCriticalityTable(XWPFDocument document, List<CriticalityTable> criticalityTableList) {
        // Find the target table (assuming it's the fourth table in the document)
        List<XWPFTable> tables = document.getTables();
        if (tables.size() >= 4) { // Ensure there are at least four tables in the document
            XWPFTable targetTable = tables.get(3); // Assuming the fourth table is the target table

            // Clear existing rows (excluding header row)
            int rowsToRemove = targetTable.getNumberOfRows() - 1;
            for (int i = 0; i < rowsToRemove; i++) {
                targetTable.removeRow(1);
            }

            // Add data to the table
            for (CriticalityTable rowData : criticalityTableList) {
                XWPFTableRow row = targetTable.createRow();
                populateCriticalityTableRow(row, rowData);
            }
        } else {
            System.out.println("Error: The document does not contain the expected number of tables.");
        }
    }

    private void populateCriticalityTableRow(XWPFTableRow row, CriticalityTable rowData) {
        List<XWPFTableCell> cells = row.getTableCells();
        if (cells.size() >= 7) { // Ensure there are at least seven cells in the row
            // Populate the cells with dynamic values
            populateCell(cells.get(0), "<AssetType>", rowData.getGroupName());
            populateCell(cells.get(1), "<Critical>", String.valueOf(rowData.getCriticalCount()));
            populateCell(cells.get(2), "<High>", String.valueOf(rowData.getHighCount()));
            populateCell(cells.get(3), "<Medium>", String.valueOf(rowData.getMediumCount()));
            populateCell(cells.get(4), "<Low>", String.valueOf(rowData.getLowCount()));
            populateCell(cells.get(5), "<Info>", String.valueOf(rowData.getInfoCount()));

            // Calculate and populate the Total cell
            int total = rowData.getCriticalCount() + rowData.getHighCount() + rowData.getMediumCount() +
                    rowData.getLowCount() + rowData.getInfoCount();
            populateCell(cells.get(6), "<Total>", String.valueOf(total));
        } else {
            System.out.println("Error: The table row does not contain the expected number of cells.");
        }
    }

    private void populateCell(XWPFTableCell cell, String placeholder, String value) {
        List<XWPFParagraph> paragraphs = cell.getParagraphs();
        if (paragraphs.isEmpty()) {
            // Clear existing content and add new text
//            cell.removeParagraph(0);
            XWPFParagraph paragraph = cell.addParagraph();
            paragraph.createRun().setText(value);
        } else {
            System.out.println("Error: The cell does not contain any paragraphs.");
        }
    }


    public static byte[] convertDocumentToByteArray(String docPath) throws IOException {
        Path path = Path.of(docPath);
        return Files.readAllBytes(path);
    }

    public static byte[] convertImageToByteArray(String imagePath) throws IOException {
        Path path = Path.of(imagePath);
        return Files.readAllBytes(path);
    }
}