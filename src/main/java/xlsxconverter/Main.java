package xlsxconverter;

import com.aspose.cells.Workbook;

import java.util.List;

public class Main {
    public static void main(String[] args) {
        XlsxConverter xlsxConverter = new XlsxConverter();
        String path = xlsxConverter.getOwnPath();
        List<Workbook> workbooks = xlsxConverter.createWorkBooks(path);
        for (Workbook actualWorkbook : workbooks) {
            xlsxConverter.setFilePathWithoutExtension(actualWorkbook.getFileName());
            xlsxConverter.deleteAdditionalSheet(actualWorkbook);
            xlsxConverter.saveXlsxAsPdf(actualWorkbook);
            xlsxConverter.saveXlsxAsModifiedXlsx(actualWorkbook);
        }
    }
}
