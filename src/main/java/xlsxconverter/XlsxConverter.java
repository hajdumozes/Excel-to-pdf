package xlsxconverter;

import com.aspose.cells.*;

import java.io.File;
import java.net.URLDecoder;
import java.nio.charset.StandardCharsets;
import java.util.*;
import java.time.LocalDate;

public class XlsxConverter {

    public static final String XLSX_EXTENSION = ".xlsx";
    public static final String PDF_EXTENSION = ".pdf";
    private String filePathWithoutExtension;
    private List<File> xlsxFiles = new ArrayList<>();
    private List<String> cellsToClear = Arrays.asList(
            "C3", "C4", "C6", "C7", "C9", "C10", "C12", "C13", "C15", "C16", "C17", "C19",
            "E3", "E4", "E6", "E7", "E9", "E10", "E12", "E13", "E15", "E16", "E17", "E19",
            "G3", "G4", "G6", "G7", "G9", "G10", "G12", "G13", "G15", "G16", "G17", "G19",
            "I3", "I4", "I6", "I7", "I9", "I10", "I12", "I13", "I15", "I16", "I17", "I19",
            "K3", "K4", "K6", "K7", "K9", "K10", "K12", "K13", "K15", "K16", "K17", "K19",
            "B21", "D21", "F21", "H21", "J21", "N8");

    public String getOwnPath() {
        ClassLoader loader = XlsxConverter.class.getClassLoader();
        String path = loader.getResource("xlsxconverter/XlsxConverter.class").getPath();
        path = path.replaceAll("%20", " ");
        String[] pathPieces = path.split("/");
        StringBuilder url = new StringBuilder();
        for (int i = 1; i < pathPieces.length - 3; i++) {
            url.append(pathPieces[i]);
            url.append(File.separator);
        }
        return URLDecoder.decode(url.toString(), StandardCharsets.UTF_8);
    }


    public void collectXlsxDirectories(String path) {
        File directory = new File(path);
        File[] files = directory.listFiles();
        if (files != null) {
            for (File file : files) {
                System.out.println(file.getAbsolutePath());
                if (file.isDirectory()) {
                    Optional<File> xlsxFile = chooseXlsxFileInDirectory(file.getPath());
                    xlsxFile.ifPresent(this::addXlsxFilesInDirectoriesToList);
                }
            }
        }
    }

    public Optional<File> chooseXlsxFileInDirectory(String path) {
        File directory = new File(path);
        File[] files = directory.listFiles();
        if (files != null) {
            for (File file : files) {
                if (file.isFile() && file.getName().substring(file.getName().indexOf(".")).equals(".xlsx")) {
                    return Optional.of(file);
                }
            }
        }
        return Optional.empty();
    }

    public void addXlsxFilesInDirectoriesToList(File xlsxFile) {
        xlsxFiles.add(xlsxFile);
    }

    public List<Workbook> createWorkBooks(String path) {
        collectXlsxDirectories(path);
        List<Workbook> workbooks = new ArrayList<>();
        for (File file : xlsxFiles) {
            try {
                workbooks.add(new Workbook(file.getPath()));
            } catch (Exception e) {
                throw new IllegalStateException("Couldn't create workbook" + e.getMessage());
            }
        }
        return workbooks;
    }

    public void setFilePathWithoutExtension(String path) {
        filePathWithoutExtension = path.substring(0, path.indexOf("."));
    }

    public void saveXlsxAsPdf(Workbook workbook) {
// Define PdfSaveOptions
        PdfSaveOptions saveOptions = new PdfSaveOptions();
// Set the compliance type
        saveOptions.setCompliance(PdfCompliance.PDF_A_1_B);
// Save the PDF file
        try {
            LocalDate localDate = LocalDate.now();
            String pathWithDirectory = createDateBasedDirectory();
            workbook.save(
                    pathWithDirectory + " - " + localDate.toString() + PDF_EXTENSION, saveOptions);
        } catch (Exception e) {
            throw new IllegalStateException("Couldn't save/overwrite PDF" + e.getMessage());
        }
    }

    public String createDateBasedDirectory() {
        String pathWithoutName = filePathWithoutExtension.substring(0, filePathWithoutExtension.lastIndexOf(File.separator));
        File directory = new File(pathWithoutName + File.separator +
                LocalDate.now().getYear() + "-" + (LocalDate.now().getYear() + 1));
        directory.mkdirs();
        return directory.getPath() + filePathWithoutExtension.substring(filePathWithoutExtension.lastIndexOf(File.separator));
    }

    public void deleteAdditionalSheet(Workbook workbook) {
        workbook.getWorksheets().removeAt(1);
    }

    public void saveXlsxAsModifiedXlsx(Workbook workbook) {
        Worksheet sheet = workbook.getWorksheets().get(0);
        Cells cells = sheet.getCells();
        clearContentsOfSelectedFields(cells);
        try {
            workbook.save(filePathWithoutExtension + XLSX_EXTENSION);
        } catch (Exception e) {
            throw new IllegalStateException("Couldn't save/overwrite XLSX" + e.getMessage());
        }
    }

    public void clearContentsOfSelectedFields(Cells cells) {
        for (String selectedCell : cellsToClear) {
            Cell actualCell = cells.get(selectedCell);
            actualCell.setValue(null);
        }
    }
}
