package to.kit.example;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidOperationException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.Stream;

public class ExampleMain {
    private void load(File file) throws IOException {
        try (Workbook workbook = WorkbookFactory.create(file, null, true)) {
            for (Sheet sheet : workbook) {
                System.out.format("\t[%s]\n", sheet.getSheetName());
                for (Row row : sheet) {
                    List<String> valueList = new ArrayList<>();
                    boolean hasValue = false;

                    for (Cell cell : row) {
                        CellType type = cell.getCellType();

                        if (type == CellType.STRING) {
                            String value = cell.getStringCellValue();

                            value = value.replace('\n', ' ');
                            valueList.add(value);
                            hasValue = true;
                        } else {
                            valueList.add("");
                        }
                    }
                    if (hasValue) {
                        System.out.format("\t%s\n", String.join("|", valueList));
                    }
                }
            }
        } catch (InvalidOperationException | EncryptedDocumentException ex) {
//            ex.printStackTrace();
        }
    }

    private void execute(List<File> dirList) throws IOException {
        for (File dir : dirList) {
            for (File file : dir.listFiles()) {
                String name = file.getName();

                if (!name.endsWith(".xls") && !name.endsWith(".xlsx")) {
                    continue;
                }
                System.out.format("'%s'\n", file.getName());
                load(file);
            }
        }
    }

    public static void main(String... args) throws Exception {
        List<File> dirList = Stream.of(args).map(File::new)
                .filter(File::exists).filter(File::isDirectory)
                .collect(Collectors.toList());
        ExampleMain app = new ExampleMain();

        app.execute(dirList);
    }
}
