package database.java.sele;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ResourceBundle;

public class DatabaseJavaexmpl {
	
	public static void main(String[] args) throws SQLException {
/*
 * create connection using connection string/stream
 * create the statement by connecting stmnt interface object with connection object
 * execute the statement with executequery method
 * verify the result/output in DB
 */
		//don't hardcode the db credentials , to hide use below format
		
		ResourceBundle rbl= ResourceBundle.getBundle("config");
		String urls=rbl.getString("url");
		String uname=rbl.getString("username");
		String pwd=rbl.getString("password");
		
		//creating connection to database
			
		Connection con=DriverManager.getConnection(urls, uname, pwd);
		
		//creating statement
		
		Statement stmnt=con.createStatement();
		
		//storing sql query in astring variable
		
		//String s="insert into employdetails values(111,'gana','inactive',32400,TO_DATE('08-03-2023','dd-mm-yyyy'))";
		//String s="update employdetails set emphiredate=TO_DATE('08-03-2023','dd-mm-yyyy') where empid=105";
		String s="delete from employdetails where emphiredate is null";
		
		//executing the sql query of string variable 
		
		stmnt.executeQuery(s);
		
		//once the query is executed close the connection
		
		con.close();
		
		System.out.println("query executed");
	}

}
