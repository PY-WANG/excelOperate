package excelOperate;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.concurrent.TimeUnit;

public class Test {
	public static void main(String[] args) {
		// DateFormat dFormat = DateFormat.getDateInstance();
		// System.out.println(dFormat.format(new Date()));
		// try {
		// DateFormat dFormat2 = new SimpleDateFormat("yyyy/MM/dd");
		// Date date = dFormat2.parse("2018/01/25");
		// System.out.println(date);
		// } catch (Exception e) {
		// // TODO: handle exception
		// }
		String fileName = "D:/Program_Files/eclipse se/workspace/excelOperate/2016-2017.xls";
		System.out.println(ExcelFile.readString(fileName, 0, 0, 0));
		// System.out.println(ExcelFile.readDate(fileName, 0, 8, 0));
		DateFormat dFormat2 = new SimpleDateFormat("yyyy/MM/dd");
		Date date = null;
		Date date1 = null;
		try {
			date = dFormat2.parse("2018/01/10");
			// date1 = dFormat2.parse();
			// date = new Date();
			// TimeUnit.MILLISECONDS.sleep(100);
			// date1 = new Date();
			// System.out.println(date.equals(date1));
			// System.out.println(dFormat2.format(date1).equals(dFormat2.format(date)));
			System.out.println(date);
			System.out.println(ExcelFile.readDate(fileName, 0, 8, 0));
			System.out.println(dFormat2.format(date).equals(dFormat2.format(ExcelFile.readDate(fileName, 0, 8, 0))));
		} catch (Exception e) {
			// TODO: handle exception
			e.printStackTrace();
		}
	}
}
