package database.java.sele;

import java.sql.Statement;
import java.awt.Color;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ResourceBundle;


import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DatabasetoExcel {
	
	public static void main(String[] args) throws SQLException, IOException {
		
		ResourceBundle rb=ResourceBundle.getBundle("config");
		String url=rb.getString("url");
		String uname=rb.getString("username");
		String pw=rb.getString("password");
		
		Connection con=DriverManager.getConnection(url, uname, pw);
		Statement st=con.createStatement();
		
		ResultSet rs=st.executeQuery("select * from customers1");
		
		XSSFWorkbook wb=new XSSFWorkbook();
		XSSFSheet sh=wb.createSheet("cust_data");
		
		
		XSSFRow row=sh.createRow(0); //creating a header row
		
		//seeting up headers nothing but 0th row
	
		
		
		row.createCell(0).setCellValue("ID");
		row.createCell(1).setCellValue("NAME");
		row.createCell(2).setCellValue("AGE");
		row.createCell(3).setCellValue("ADDRESS");
		row.createCell(4).setCellValue("SALARY");
		
		int r=1;
		while(rs.next()) { // iterate thru each record (cell by cell) and store the cell values
			
			
			
			int empid=rs.getInt("ID");
			String empname=rs.getString("NAME");
			int empage=rs.getInt("AGE");
			String empaddr=rs.getString("ADDRESS");
			int empsal=rs.getInt("SALARY");
			
			row=sh.createRow(r);  //create 1st row in xcel and above data into each cell 
			r++;
			row.createCell(0).setCellValue(empid);
			row.createCell(1).setCellValue(empname);
			row.createCell(2).setCellValue(empage);
			row.createCell(3).setCellValue(empaddr);
			row.createCell(4).setCellValue(empsal);
			
		
			
		
		}
		FileOutputStream fo=new FileOutputStream(".\\src\\test\\resources\\dbtoexcel.xlsx");
	
			wb.write(fo);
		
		
		
		wb.close();
		fo.close();
		con.close();
		
		
	}

}
