/**
 * 
 */
package com.excel;

import java.io.IOException;
import java.util.List;


/**
 * @author Hongten
 * @created 2014-5-21
 */
public class Client {

    public static void main(String[] args) throws IOException {
        String excel2003_2007 = Common.STUDENT_INFO_XLS_PATH;
        //String excel2010 = Common.STUDENT_INFO_XLSX_PATH;
    	//String excel2003_2007 = "C:/Users/Administrator.LDNHHM-PC/Desktop/test3.xls";
        String excel2010 = "C:/Users/Administrator.LDNHHM-PC/Desktop/aaa.xlsx";
        // read the 2003-2007 excel
        try {
			
        	List<Student> list = new ReadExcel().readExcel(excel2003_2007);
        	if (list != null) {
        		for (Student student : list) {
        			System.out.println("No. : " + student.getNo() + ", name : " + student.getName() + ", age : " + student.getAge() + ", score : " + student.getScore());
        		}
        	}
        	System.out.println("======================================");
        	// read the 2010 excel
        	List<Student> list1 = new ReadExcel().readExcel(excel2010);
        	if (list1 != null) {
        		for (Student student : list1) {
        			System.out.println("No. : " + student.getNo() + ", name : " + student.getName() + ", age : " + student.getAge() + ", score : " + student.getScore());
        		}
        	}
		} catch (Exception e) {
			e.printStackTrace();
		}
    }
}