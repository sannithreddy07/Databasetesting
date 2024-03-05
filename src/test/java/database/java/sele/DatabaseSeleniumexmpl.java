package database.java.sele;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.time.Duration;
import java.util.ResourceBundle;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;


public class DatabaseSeleniumexmpl {
	
	public static void main(String[] args) throws SQLException {
		/*
		 * To get data from database table and test the data in the application using Selenium
		 */
		WebDriver driver;
		
		ResourceBundle rbl=ResourceBundle.getBundle("config");
		String urls=rbl.getString("url");
		String uname=rbl.getString("username");
		String pwd=rbl.getString("password");
		
		
		Connection con=DriverManager.getConnection(urls, uname, pwd);
		
		Statement stm=con.createStatement();
		
		String s1="select * from logindetails"; 
		
		ResultSet rs=stm.executeQuery(s1); //storing the results of executed query in Resutset variable to view results in console
		
		while(rs.next()) { // to read each record from results in table
			//to get each cell data use below method , i used index here we can also use coloumn name as well
			String uname1=rs.getString(1);
			String pswd=rs.getString(2);

			driver=new ChromeDriver(); //for each record open browser and enter url and validate credentials
			driver.manage().window().maximize();
			driver.navigate().to(rbl.getString("urll"));
			driver.findElement(By.name("txtUserName")).sendKeys(uname1);
			driver.findElement(By.name("txtPassword")).sendKeys(pswd);
			driver.findElement(By.name("Submit")).click();
			driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(3));
			if(driver.getTitle().equals("OrangeHRM")) {
				driver.findElement(By.linkText("Logout"));
				System.out.println("test case passed");
			}else {
				System.out.println("test case failed");
			}
			driver.close();
			
		}
		
		rs.close(); //close the resultset
		stm.close(); //close the statement connection
		con.close(); //close the jdbc connection
		System.out.println("query executed");
	}

}
