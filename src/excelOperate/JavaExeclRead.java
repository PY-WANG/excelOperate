package excelOperate;

import java.io.*;
import java.io.File;

import jxl.*;
import jxl.read.biff.*;

public class JavaExeclRead {
	public static void main(String[] args) {
		try {
			Workbook workbook = Workbook
					.getWorkbook(new File("D:/Program_Files/eclipse se/workspace/excelOperate/2016-2017.xls"));
			Sheet sheet1 = workbook.getSheet(0);
			Cell colArow1 = sheet1.getCell(0, 0);
			Cell colArow4 = sheet1.getCell(0, 4);
			NumberCell numberCell = (NumberCell) colArow4;
			LabelCell labelCell = (LabelCell) colArow1;
			System.out.println(labelCell.getString());
			System.out.println(numberCell.getValue());
		} catch (BiffException e) {
			// TODO: handle exception
			e.printStackTrace();
		} catch (IOException e) {
			// TODO: handle exception
			e.printStackTrace();
		}
	}
}
