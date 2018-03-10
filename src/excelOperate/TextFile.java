package excelOperate;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.io.PrintWriter;

public class TextFile {
	public static String read(String fileName) {
		StringBuilder sb = new StringBuilder();
		try {
			BufferedReader in = new BufferedReader(new FileReader(new File(fileName).getAbsolutePath()));
			try {
				String s;
				while ((s = in.readLine()) != null) {
					sb.append(s);
					sb.append("\n");
				}
			} finally {
				in.close();
			}
		} catch (IOException e) {
			// TODO: handle exception
			throw new RuntimeException(e);
		}
		return sb.toString();
	}

	public static void write(String fileName, String text) {
		try {
			PrintWriter out = new PrintWriter(new File(fileName).getAbsolutePath());
			try {
				out.print(text);
			} finally {
				out.close();
			}
		} catch (IOException e) {
			// TODO: handle exception
			throw new RuntimeException(e);
		}
	}
	public static void main(String[] args){
		String letterMode = read("D:/Program_Files/eclipse se/workspace/excelOperate/letterMode.xml");
		System.out.println(letterMode);
	}
}
