package excelOperate;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.io.PrintWriter;

public class TxtRead {
	public static void main(String[] args) {
		// StringBuilder sb = new StringBuilder();
		// try {
		// BufferedReader in = new BufferedReader(new FileReader(
		// new File("D:/Program_Files/eclipse
		// se/workspace/excelOperate/letter.txt").getAbsolutePath()));
		// try {
		// String s;
		// while ((s = in.readLine()) != null) {
		// sb.append(s);
		// sb.append("\n");
		// }
		// } finally {
		// in.close();
		// }
		// } catch (IOException e) {
		// throw new RuntimeException(e);
		// // TODO: handle exception
		// }
		// System.out.println(sb.toString());
		try {
			PrintWriter out = new PrintWriter(
					new File("D:/Program_Files/eclipse se/workspace/excelOperate/letter1.txt").getAbsolutePath());
			try {
				out.print("asd);lfkjasd;lkfjasl;dk;asdkfja;lsd");
			} finally {
				out.close();
			}
		} catch (IOException e) {
			// TODO: handle exception
			throw new RuntimeException(e);
		}
	}
}
