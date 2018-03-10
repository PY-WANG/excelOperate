package excelOperate;

import java.util.ArrayList;
import java.util.Iterator;

public class StringReplace {
	static String letterName = "D:/CODE/Documents/letterMode.xml";
	static String fileName = "D:/CODE//Documents/2016-2017.xls";
	static String dateString = "2018-03-03";
	static int dateCol = ExcelFile.getDateCol(dateString, fileName);
	static String[] repedString = { "%Author%", "%PaperTitle%", "%volumeNum%", "%issueNum%", "%quotedNum%", "%Email%" };
	static String letter = TextFile.read(letterName);

	/**
	 * get the content of a rowContent
	 * 
	 * @param rowContent
	 * @return
	 */
	public String[] getRepString(RowContent rowContent) {
		String[] repString = new String[6];
		repString[0] = rowContent.getAuthor();
		repString[1] = rowContent.getPaperTitle();
		repString[2] = rowContent.getVolumeNum() + "";
		repString[3] = rowContent.getIssueNum() + "";
		repString[4] = rowContent.getQuotedNum() + "";
		repString[5] = rowContent.getEmail();
		return repString;
	}

	/**
	 * using replaceAll() replace keywords of letter and using Recursive
	 * 
	 * @param i
	 * @param rowContent
	 * @return
	 */
	public String replace(int i, RowContent rowContent) {
		if (i == 0) {
			return letter;
		} else {
			return replace(i - 1, rowContent).replaceAll(repedString[i - 1], getRepString(rowContent)[i - 1]);
		}
	}

	/**
	 * save these letters by merge
	 * 
	 * @param lettered
	 */
	public void saveByMerge(ArrayList<String> lettered) {
		StringBuilder out = new StringBuilder();
		Iterator<String> istring = lettered.iterator();
		while (istring.hasNext()) {
			out.append(istring.next());
		}
		System.out.println(out);
		TextFile.write("D:/CODE/Documents/letter/out-" + dateString + ".doc",
				out.toString());
	}
	
	/**
	 * save these letters by single
	 * 
	 * @param lettered
	 */
	public void saveBySingle(ArrayList<String> lettered) {
		Iterator<String> istring = lettered.iterator();
		int i = 1;
		while (istring.hasNext()) {
			String letter = istring.next();
			System.out.println(letter);
			TextFile.write("D:/CODE/Documents/letter/out-" + dateString + " - " + i + ".doc",
					letter.toString());
			i++;
		}
	}

	public static void main(String[] args) {
		ArrayList<RowContent> rowContents = ExcelFile.getAllRowContent(fileName, dateCol);
		Iterator<RowContent> it = rowContents.iterator();
		ArrayList<String> lettered = new ArrayList<>();
		StringReplace stringReplace = new StringReplace();
		while (it.hasNext()) {
			lettered.add(stringReplace.replace(6, it.next()));
		}
		// 合并输出
		//stringReplace.saveByMerge(lettered);
		stringReplace.saveBySingle(lettered);
	}
}
