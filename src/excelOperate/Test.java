package excelOperate;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.concurrent.TimeUnit;

public class Test {
	
	static final String JDBC_DRIVER = "com.mysql.jdbc.Driver";  
    static final String DB_URL = "jdbc:mysql://localhost:3306/paperdb";
 
    // 数据库的用户名与密码，需要根据自己的设置
    static final String USER = "root";
    static final String PASS = "123456";
    
    static String fileName = "D:/CODE//Documents/2016-2017.xls";
	
	public static void main(String[] args) {
		// 存数据
		ArrayList<RowContent> rowContents = ExcelFile.getAllRowContent(fileName, StringReplace.dateCol);
		Connection conn = null;
		Statement stmt = null;
		PreparedStatement preparedStatement = null;// 预处理
		
		try {
			Class.forName(JDBC_DRIVER);
			System.out.println("连接数据库");
			conn = DriverManager.getConnection(DB_URL, USER, PASS);
			
			//执行查询
			System.out.println("实例化Statement对象");
			stmt = conn.createStatement();
			String sql = "Select * from paper_quoted_flow";
			ResultSet rs = stmt.executeQuery(sql);
			
			while(rs.next()) {
				int flowId = rs.getInt("flow_id");
				int paperId = rs.getInt("paper_id");
				int quotedNum = rs.getInt("quoted_num");
				Date date = rs.getDate("create_datetime");
				System.out.println("flowId:" + flowId);
				System.out.println("paperId:" + paperId);
				System.out.println("quotedNum:" + quotedNum);
				System.out.println("date:" + date);
			}
			
			// 插入数据
			Iterator<RowContent> it = rowContents.iterator();
			while(it.hasNext()) {
				RowContent temp = it.next();
				try {
					String insertInfo = "insert into paper_info (paper_title, author, email, volume_num, issue_num) values(?,?,?,?,?)";
					preparedStatement = conn.prepareStatement(insertInfo);
					preparedStatement.setString(1, temp.getPaperTitle());
					preparedStatement.setString(2, temp.getAuthor());
					preparedStatement.setString(3, temp.getEmail());
					preparedStatement.setInt(4, temp.getVolumeNum());
					preparedStatement.setInt(5, temp.getIssueNum());
					preparedStatement.executeUpdate();
				} catch(SQLException e) {
					e.printStackTrace();
				}
			}
			
			rs.close();
			stmt.close();
			conn.close();
		}catch(SQLException se){
            // 处理 JDBC 错误
            se.printStackTrace();
        }catch(Exception e){
            // 处理 Class.forName 错误
            e.printStackTrace();
        }finally{
            // 关闭资源
            try{
                if(stmt!=null) stmt.close();
            }catch(SQLException se2){
            }// 什么都不做
            try{
                if(conn!=null) conn.close();
            }catch(SQLException se){
                se.printStackTrace();
            }
        }
/*		
		System.out.println(ExcelFile.readString(fileName, 0, 3, 3));
		// System.out.println(ExcelFile.readDate(fileName, 0, 8, 0));
		DateFormat dFormat2 = new SimpleDateFormat("yyyy/MM/dd");
		Date date = null;
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
		}*/
	}
}
