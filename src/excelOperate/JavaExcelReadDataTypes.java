package excelOperate;

import java.io.File;
import java.io.IOException;

import jxl.BooleanCell;
import jxl.Cell;
import jxl.CellType;
import jxl.DateCell;
import jxl.LabelCell;
import jxl.NumberCell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class JavaExcelReadDataTypes {
	public static void main(String[] args) {
		Workbook workbook;
		try {
			workbook = Workbook.getWorkbook(new File("D:/Program_Files/eclipse se/workspace/excelOperate/test.xls"));
			Sheet sheet1 = workbook.getSheet(0);
			// obtain reference to the cell
			Cell cell1 = sheet1.getCell(0, 0);
			Cell cell2 = sheet1.getCell(1, 0);
			Cell cell3 = sheet1.getCell(2, 0);
			Cell cell4 = sheet1.getCell(3, 0);
			DateCell dateCell = null;
			NumberCell numberCell = null;
			BooleanCell booleanCell = null;
			LabelCell labelCell = null;
			// check the type of the cell contents
			if (cell1.getType() == CellType.DATE)
				dateCell = (DateCell) cell1;
			if (cell2.getType() == CellType.NUMBER)
				numberCell = (NumberCell) cell2;
			if (cell3.getType() == CellType.BOOLEAN)
				booleanCell = (BooleanCell) cell3;
			if (cell4.getType() == CellType.LABEL)
				labelCell = (LabelCell) cell4;
			// display the contents of the cell
			System.out.println("Value of Date Cell: " + dateCell.getDate());
			System.out.println("Value of Number Cell: " + numberCell.getValue());
			System.out.println("Value of Boolean Cell: " + booleanCell.getValue());
			System.out.println("Value of Label Cell: " + labelCell.getString());
		} catch (BiffException e) {
			e.printStackTrace();
			// TODO: handle exception
		} catch (IOException e) {
			e.printStackTrace();
			// TODO: handle exception
		}
	}
}
