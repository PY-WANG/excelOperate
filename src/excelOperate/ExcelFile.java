package excelOperate;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;

import jxl.Cell;
import jxl.CellType;
import jxl.DateCell;
import jxl.LabelCell;
import jxl.NumberCell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.read.biff.BiffException;

public class ExcelFile {
	/**
	 * read a cell
	 * 
	 * @param fileName
	 * @param sheetNum
	 * @param col
	 * @param row
	 * @return a Cell
	 */
	public static Cell readCell(String fileName, int sheetNum, int col, int row) {
		WorkbookSettings settings = new WorkbookSettings();
		settings.setEncoding("UTF-8");
		Workbook workbook;
		Cell cell = null;
		try {
			InputStream is = new FileInputStream((new File(fileName)).getAbsolutePath());
			workbook = Workbook.getWorkbook(is, settings);
			Sheet sheet = workbook.getSheet(sheetNum);
			if (col >= sheet.getColumns() || row >= sheet.getRows()) {
				return null;
			}
			cell = sheet.getCell(col, row);
			workbook.close();
		} catch (BiffException e) {
			// TODO: handle exception
			e.printStackTrace();
		} catch (IOException e) {
			// TODO: handle exception
			e.printStackTrace();
		}
		return cell;
	}

	/**
	 * return the row number of a sheet in a file
	 * 
	 * @param fileName
	 * @param sheetNum
	 * @return
	 */
	public static int getRowNum(String fileName, int sheetNum) {
		Workbook workbook;
		int rowNum = -1;
		try {
			InputStream is = new FileInputStream((new File(fileName)).getAbsolutePath());
			workbook = Workbook.getWorkbook(is);
			Sheet sheet = workbook.getSheet(sheetNum);
			rowNum = sheet.getRows();
		} catch (FileNotFoundException e) {
			// TODO: handle exception
			e.printStackTrace();
		} catch (BiffException e) {
			// TODO: handle exception
			e.printStackTrace();
		} catch (IOException e) {
			// TODO: handle exception
			e.printStackTrace();
		}
		return rowNum;
	}

	/**
	 * return the row number of the first sheet in a file
	 * 
	 * @param fileName
	 * @return
	 */
	public static int getRowNum(String fileName) {
		return getRowNum(fileName, 0);
	}

	/**
	 * read a cell of Date Type
	 * 
	 * @param cell
	 * @return
	 */
	public static Date readDate(Cell cell) {
		DateCell dateCell = null;
		if (cell.getType() == CellType.DATE) {
			dateCell = (DateCell) cell;
			return dateCell.getDate();
		} else {
			return null;
		}
	}

	/**
	 * read a cell of Date Type
	 * 
	 * @param fileName
	 * @param sheetNum
	 * @param col
	 * @param row
	 * @return Date
	 */
	public static Date readDate(String fileName, int sheetNum, int col, int row) {
		Cell cell = readCell(fileName, sheetNum, col, row);
		return readDate(cell);
	}

	/**
	 * read a cell of int Type
	 * 
	 * @param cell
	 * @return
	 */
	public static int readInt(Cell cell) {
		NumberCell numberCell = null;
		if (cell.getType() == CellType.NUMBER) {
			numberCell = (NumberCell) cell;
			return (int) numberCell.getValue();
		} else {
			return -1;
		}
	}

	/**
	 * read a cell of int Type
	 * 
	 * @param fileName
	 * @param sheetNum
	 * @param col
	 * @param row
	 * @return
	 */
	public static int readInt(String fileName, int sheetNum, int col, int row) {
		Cell cell = readCell(fileName, sheetNum, col, row);
		return readInt(cell);
	}

	/**
	 * read a cell of String Type
	 * 
	 * @param cell
	 * @return
	 */
	public static String readString(Cell cell) {
		LabelCell labelCell = null;
		if (cell.getType() == CellType.LABEL) {
			labelCell = (LabelCell) cell;
			String s = labelCell.getString().trim();
			// System.out.println(s);
			// byte bytes[] = { (byte) 0xC2, (byte) 0xA0 };
			// String UTFSpace = new String(bytes, "utf-8");
			return s;
		} else {
			return null;
		}
	}

	/**
	 * read a cell of String Type
	 * 
	 * @param fileName
	 * @param sheetNum
	 * @param col
	 * @param row
	 * @return
	 */
	public static String readString(String fileName, int sheetNum, int col, int row) {
		Cell cell = readCell(fileName, sheetNum, col, row);
		return readString(cell);
	}

	/**
	 * get the colume of special date(if exist,return the number; if not exist,
	 * return -1)
	 * 
	 * @param dateString
	 * @param fileName
	 * @return
	 */
	public static int getDateCol(String dateString, String fileName) {
		int dateCol;
		DateFormat dFormat = new SimpleDateFormat("yyyy-MM-dd");
		Date date = null;
		try {
			date = dFormat.parse(dateString);
		} catch (Exception e) {
			// TODO: handle exception
			e.printStackTrace();
		}
		// System.out.println(date);
		Cell cell = null;
		int colnum = 0;
		while (true) {
			cell = readCell(fileName, 0, colnum, 0);
			// System.out.println(colnum);
			if (cell == null) {
				dateCol = -1;
				break;
			} else if (cell.getType() != CellType.DATE) {
				colnum++;
				// System.out.println("not equal");
				continue;
			} else if (cell.getType() == CellType.DATE) {
				Date dateColnum = readDate(cell);
				// System.out.println("equal");
				if (dFormat.format(date).equals(dFormat.format(dateColnum))) {
					dateCol = colnum;
					// System.out.println("break");
					break;
				}
				colnum++;
			}
		}
		return dateCol;
	}

	/**
	 * read a special row
	 * 
	 * @param fileName
	 * @param sheetNum
	 * @param row
	 * @param dateCol
	 * @return
	 */
	public static RowContent getRowContent(String fileName, int sheetNum, int row, int dateCol) {
		if (row < 1) {
			return null;
		}
		RowContent rowContent = new RowContent();
		rowContent.setAuthor(readString(fileName, sheetNum, 4, row));
		rowContent.setPaperTitle(readString(fileName, sheetNum, 3, row));
		rowContent.setVolumeNum(readInt(fileName, sheetNum, 1, row));
		rowContent.setIssueNum(readInt(fileName, sheetNum, 2, row));
		rowContent.setQuotedNum(readInt(fileName, sheetNum, dateCol, row));
		rowContent.setEmail(readString(fileName, sheetNum, 6, row));
		return rowContent;
	}

	/**
	 * read a special row
	 * 
	 * @param fileName
	 * @param row
	 * @param dateCol
	 * @return
	 */
	public static RowContent getRowContent(String fileName, int row, int dateCol) {
		return getRowContent(fileName, 0, row, dateCol);
	}

	public static ArrayList<RowContent> getAllRowContent(String fileName, int dateCol) {
		ArrayList<RowContent> rowContents = new ArrayList<>();
		int rowNum = getRowNum(fileName);
		for (int i = 1; i < rowNum; i++) {
			RowContent rContent = getRowContent(fileName, i, dateCol);
			// System.out.println(rContent);
			rowContents.add(rContent);
		}
		return rowContents;
	}

	public static void main(String[] args) {
		String fileName = "D:/Program_Files/eclipse se/workspace/excelOperate/2016-2017.xls";
		String dateString = "2018-01-24";
		int dateCol = getDateCol(dateString, fileName);
		System.out.println(dateCol);
		// System.out.println(readDate(fileName, 0, 9, 0));
		for (int i = 1; i < 2; i++)
			System.out.println(getRowContent(fileName, i, dateCol));
		// ArrayList<RowContent> rowContents = getAllRowContent(fileName,
		// dateCol);
		// Iterator<RowContent> iterator = rowContents.iterator();
		// while (iterator.hasNext()) {
		// System.out.println(iterator.next());
		// }
	}
}
