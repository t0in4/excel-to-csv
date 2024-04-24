import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintStream;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {
	@SuppressWarnings("rawtypes")
	public static void main(String[] args) throws InvalidFormatException, IOException {
		XSSFWorkbook wb = new XSSFWorkbook(new File("file.xlsx")); // in root folder of project absolute path
		DataFormatter formatter = new DataFormatter();
		PrintStream out = new PrintStream(new FileOutputStream("file.csv"), true, "UTF-8"); // in root folder of project absolute path
		for (Sheet sheet : wb) {
			for (Row row : sheet) {
				boolean firstCell = true;
				for (Cell cell : row) {
					if (!firstCell)
						out.print(',');
					String text = formatter.formatCellValue(cell);
					out.print(text);
					firstCell = false;
				}
				out.println();
			}
		}
	}

}
