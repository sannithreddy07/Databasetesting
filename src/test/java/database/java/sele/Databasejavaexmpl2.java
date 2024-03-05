package database.java.sele;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ResourceBundle;

public class Databasejavaexmpl2 {
public static void main(String[] args) throws SQLException {
	/*
	 * To see the results of DB in console for query execution use select 
	 * select is present in ResultSet Interface 
	 * To store the result of executed query we ResultSet reference variable/object
	 */
	
	ResourceBundle rbl=ResourceBundle.getBundle("config");
	String urls=rbl.getString("url");
	String uname=rbl.getString("username");
	String pwd=rbl.getString("password");
	
	Connection con=DriverManager.getConnection(urls, uname, pwd);
	Statement stm=con.createStatement();
	String s1="select * from employdetails"; 
	ResultSet rs=stm.executeQuery(s1); //storing the results of executed query in Resutset variable to view results in console
	while(rs.next()) { // to read each record from results in table
		//to get each cell data use below method , i used index here we can also use coloumn name as well
		System.out.print(rs.getString(1)+" "+rs.getString(2)+" "+rs.getString(3)+" "+rs.getString(4)+" "+rs.getString(5));
		System.out.println(); //after each record move the cursor next line , so that other record will print in next line 
	}
	rs.close(); //close the resultset
	stm.close(); //close the statement connection
	con.close(); //close the jdbc connection
	System.out.println("query executed");
	}
}
