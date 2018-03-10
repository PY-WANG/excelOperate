package excelOperate;

import java.io.File;
import java.io.IOException;
import java.util.Date;

import jxl.*;
import jxl.read.biff.BiffException;
import jxl.write.*;
import jxl.write.Boolean;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.biff.RowsExceededException;

public class JavaExcelWrite {
	public static void main(String[] args) {
		File exlFile = new File("D:/Program_Files/eclipse se/workspace/excelOperate/write_test.xls");
		try {

			WritableWorkbook writableWorkbook = Workbook.createWorkbook(exlFile);

			WritableSheet writableSheet = writableWorkbook.createSheet("Sheet1", 0);

			DateFormat dFormat2 = new DateFormat("yyyy/MM/dd");
			WritableCellFormat cellDateFormat = new WritableCellFormat(dFormat2);
			Label label = new Label(0, 0, "Label(String)");
			DateTime date = new DateTime(1, 0, new Date(), cellDateFormat);
			Boolean bool = new Boolean(2, 0, true);
			Number num = new Number(3, 0, 9.99);

			writableSheet.addCell(label);
			writableSheet.addCell(date);
			writableSheet.addCell(bool);
			writableSheet.addCell(num);

			writableWorkbook.write();
			writableWorkbook.close();

		} catch (IOException e) {
			// TODO: handle exception
			e.printStackTrace();
		} catch (RowsExceededException e) {
			// TODO: handle exception
			e.printStackTrace();
		} catch (WriteException e) {
			// TODO: handle exception
			e.printStackTrace();
		} catch (Exception e) {
			// TODO: handle exception
		}
		System.out.println("Done!");
		try {
			Workbook workbook = Workbook.getWorkbook(exlFile);
			Sheet sheet = workbook.getSheet(0);
			Cell cell = sheet.getCell(1, 0);
			DateCell dCell = (DateCell) cell;
			System.out.println(dCell.getDate());
		} catch (BiffException e) {
			// TODO: handle exception
			e.printStackTrace();
		} catch (IOException e) {
			// TODO: handle exception
			e.printStackTrace();
		}
	}
}
