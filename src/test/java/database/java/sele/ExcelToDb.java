package database.java.sele;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ResourceBundle;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelToDb {
	
	
		
		public static void main(String[] args) throws SQLException, IOException {
			
			ResourceBundle rb=ResourceBundle.getBundle("config");
			String url=rb.getString("url");
			String uname=rb.getString("username");
			String pw=rb.getString("password");
			
			Connection con=DriverManager.getConnection(url, uname, pw);
			
			Statement st=con.createStatement();
			
			String query=null;
//			"CREATE TABLE CUSTOMERS3 ( ID INT NOT NULL,NAME VARCHAR (20) NOT NULL,AGE INT NOT NULL, ADDRESS VARCHAR (25),SALARY FLOAT (10),PRIMARY KEY (ID))";
//			"DELETE FROM CUSTOMERS3 WHERE ID IN (2000,1500)"
			st.executeQuery(query);
			System.out.println("table created");
			
			FileInputStream f=new FileInputStream(".\\src\\test\\resources\\dbtoexcel.xlsx");
			XSSFWorkbook wb=new XSSFWorkbook(f);
			XSSFSheet sh=wb.getSheet("cust_data");
			int rows=sh.getLastRowNum();
			
			
			for(int i=1;i<=rows;i++) {
				XSSFRow row=sh.getRow(i);
				
				int empid=(int)row.getCell(0).getNumericCellValue();
				String empname=row.getCell(1).getStringCellValue();
				int age=(int)row.getCell(2).getNumericCellValue();
				String addr=row.getCell(3).getStringCellValue();
				float sal=(float)row.getCell(4).getNumericCellValue();
				
				query="INSERT INTO CUSTOMERS3 VALUES('"+empid+"','"+empname+"', '"+age+"', '"+addr+"', '"+sal+"')";
				st.executeQuery(query);
				st.executeQuery("commit"); //to make the data permanent in table 
				
			}
			
			
			
			
			wb.close();
			f.close();
			con.close();
			
			System.out.println("values are inserted");
			
			
	}

}
