package test;

import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel {
	public static void main(String[] args) throws IOException  {





		List<String> ParentTable  = new ArrayList<String>();
		List<Integer> Age  = new ArrayList<Integer>();
		List<String> ageTable = new ArrayList<String>();
		List<String> gender = new ArrayList<String>(); 




		//	 get the value Parent.xlsx Table

		String fileLocation1 = "./BZA02/Parent Table.xlsx";
		XSSFWorkbook wbook1 = null;
		try {
			wbook1 = new XSSFWorkbook(fileLocation1);
		} catch (IOException e) {

			e.printStackTrace();
		}
		XSSFSheet sheet1 = wbook1.getSheetAt(0);

		int lastRowNum1 = sheet1.getLastRowNum();

		//int physicalNumberOfRows1 = sheet1.getPhysicalNumberOfRows();



		short lastCellNum1 = sheet1.getRow(0).getLastCellNum();

		for (int i = 1; i <=lastRowNum1; i++) {
			XSSFRow row = sheet1.getRow(i);
			for (int j = 0; j < lastCellNum1; j++) {
				XSSFCell cell = row.getCell(j);
				DataFormatter dft = new DataFormatter();
				String value = dft.formatCellValue(cell);

				ParentTable.add(value);

			} 
		}
		//		 get the value Parent Table.xlsx in Age

		for (int i = 1; i <=lastRowNum1; i++) {
			XSSFRow row = sheet1.getRow(i);
			for (int j = 2; j <lastCellNum1-1; j++) {
				XSSFCell cell = row.getCell(j);
				DataFormatter dft = new DataFormatter();
				String value = dft.formatCellValue(cell);

				int age=Integer.parseInt(value);

				Age.add(age);

			} 
		}




		//get the value Age Table.xlsx

		String fileLocation2 = "./BZA02/Age Table.xlsx";
		XSSFWorkbook wbook2 = null;
		try {
			wbook2 = new XSSFWorkbook(fileLocation2);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		XSSFSheet sheet2 = wbook2.getSheetAt(0);

		int lastRowNum2 = sheet2.getLastRowNum();

		short lastCellNum2 = sheet2.getRow(0).getLastCellNum();

		for (int i = 1; i <=lastRowNum2; i++) {
			XSSFRow row = sheet2.getRow(i);
			for (int j = 0; j < lastCellNum2; j++) {
				XSSFCell cell = row.getCell(j);
				DataFormatter dft = new DataFormatter();
				String value = dft.formatCellValue(cell);

				ageTable.add(value);
			} 
		}



		//get the value Gender Table.xlsx

		String fileLocation3 = "./BZA02/Gender Table.xlsx";
		XSSFWorkbook wbook3 = null;
		try {
			wbook3 = new XSSFWorkbook(fileLocation3);
		} catch (IOException e) {

			e.printStackTrace();
		}
		XSSFSheet sheet3 = wbook3.getSheetAt(0);

		int lastRowNum3 = sheet3.getLastRowNum();

		short lastCellNum3 = sheet3.getRow(0).getLastCellNum();

		for (int i = 1; i <=lastRowNum3; i++) {
			XSSFRow row = sheet3.getRow(i);
			for (int j = 0; j < lastCellNum3; j++) {
				XSSFCell cell = row.getCell(j);
				DataFormatter dft = new DataFormatter();
				String value = dft.formatCellValue(cell);

				gender.add(value);
			} 
		}



		/* 
		 * Seeta
		 */

		if(Age.get(0)==45 && ageTable.get(2).equals("26-50")) {

			if(gender.get(2).equals(ParentTable.get(3))) {

				String str1 =ageTable.get(3);
				String value1 = str1.substring(0, str1.length()-1);
				String str2=  gender.get(3);
				String value2 = str2.substring(0, str2.length()-1);
				int num1= Integer.parseInt(value1);
				int num2= Integer.parseInt(value2);



				int total=num1+num2;

				System.out.println("\"Id\" : "+ParentTable.get(0));
				System.out.println("\"Name\" : "+ParentTable.get(1));
				System.out.println("\"TotalConcession\" : "+total);
				//System.out.println(total);


			}	
			else System.out.println("Seeta inside error");


		}	
		else System.out.println("Seeta error");


		/*
		 * Maala
		 */

		if(Age.get(1)==24 && ageTable.get(0).equals("0-25")) {

			if(gender.get(2).equals(ParentTable.get(7))) {

				String str1 =ageTable.get(1);
				String value1 = str1.substring(0, str1.length()-1);
				String str2=  gender.get(3);
				String value2 = str2.substring(0, str2.length()-1);
				int num1= Integer.parseInt(value1);
				int num2= Integer.parseInt(value2);

				int total=num1+num2;

				System.out.println("\"Id\" : "+ParentTable.get(4));
				System.out.println("\"Name\" : "+ParentTable.get(5));
				System.out.println("\"TotalConcession\" : "+total);
				//System.out.println(total);


			}	
			else System.out.println("Maala inside error");

		}

		else System.out.println("Maala error");


		/*
		 * Ram
		 */
		if(Age.get(2)==16 && ageTable.get(0).equals("0-25")) {





			if(gender.get(0).equals(ParentTable.get(11))) {

				String str1 =ageTable.get(1);
				String value1 = str1.substring(0, str1.length()-1);
				String str2=  gender.get(1);
				String value2 = str2.substring(0, str2.length()-1);
				int num1= Integer.parseInt(value1);
				int num2= Integer.parseInt(value2);

				int total=num1+num2;

				System.out.println("\"Id\" : "+ParentTable.get(8));
				System.out.println("\"Name\" : "+ParentTable.get(9));
				System.out.println("\"TotalConcession\" : "+total);
				//System.out.println(total);



			}	
			else System.out.println("Ram inside error");

		}
		else System.out.println(" Ram error");


		/*
		 * Ajay
		 */

		if(Age.get(3)==78 && ageTable.get(6).equals("76-100")) {





			if(gender.get(0).equals(ParentTable.get(15))) {

				String str1 =ageTable.get(7);
				String value1 = str1.substring(0, str1.length()-1);
				String str2=  gender.get(1);
				String value2 = str2.substring(0, str2.length()-1);
				int num1= Integer.parseInt(value1);
				int num2= Integer.parseInt(value2);

				int total=num1+num2;

				System.out.println("\"Id\" : "+ParentTable.get(12));
				System.out.println("\"Name\" : "+ParentTable.get(13));
				System.out.println("\"TotalConcession\" : "+total);


			}	
			else System.out.println("Ajay inside error");

		}
		else System.out.println("Ajay error");


		/* 
		 * Nayak
		 *  */

		if(Age.get(4)==67 && ageTable.get(4).equals("51-75")) {




			if(gender.get(0).equals(ParentTable.get(19))) {

				String str1 =ageTable.get(5);
				String value1 = str1.substring(0, str1.length()-1);
				String str2=  gender.get(1);
				String value2 = str2.substring(0, str2.length()-1);
				int num1= Integer.parseInt(value1);
				int num2= Integer.parseInt(value2);
				int total=num1+num2;


				System.out.println("\"Id\" : "+ParentTable.get(16));
				System.out.println("\"Name\" : "+ParentTable.get(17));
				System.out.println("\"TotalConcession\" : "+total);
				//System.out.println("Id: "+total+"%");


			}	
			else System.out.println("Nayak inside error");

		}
		else System.out.println("Nayak error");



		try {
			wbook1.close();
		} catch (IOException e) {

			e.printStackTrace();
		}




	} 

}
