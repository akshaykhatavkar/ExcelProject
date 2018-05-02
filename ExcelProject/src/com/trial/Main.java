package com.trial;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.examples.CellStyleDetails;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFBorderFormatting;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {
	

	public static void main(String[] args) throws Exception {
		
		/*File f = new File("Hello.xlsx");
		FileInputStream fis = new FileInputStream(f);
		
		Workbook wb = new XSSFWorkbook(fis);
		Sheet sheet = wb.getSheet("Sheet1");
		
		System.out.println(getCellVal(sheet.getRow(1), 0));
		System.out.println(getCellVal(sheet.getRow(1), 1));
		
		sheet.getRow(1).getCell(0).setCellValue(456789);
		sheet.createRow(1);
		sheet.getRow(1).createCell(0).setCellValue(123456);
		sheet.getRow(1).createCell(1).setCellValue("Akshay K");
		
//		wb.createSheet("Data");
		Sheet sheet1 = wb.getSheet("Data");
		sheet1.createRow(0);
		sheet1.getRow(0).createCell(0).setCellValue("Hello");
//		sheet1.getRow(0).getCell(0).setCellStyle();
		sheet1.getRow(0).createCell(1).setCellValue("I am ");
		
		CellStyle style = wb.createCellStyle();
		style.setFillBackgroundColor(IndexedColors.AQUA.getIndex());
		sheet1.getRow(0).createCell(2).setCellValue("Trying something new");
		sheet1.getRow(0).getCell(0).setCellStyle(style);
		
		fis.close();
		FileOutputStream fos = new FileOutputStream(f);
		wb.write(fos);
//		fos.flush();
		fos.close();
		wb.close();*/
		
		Workbook wb = new XSSFWorkbook();
		wb.createSheet("Data");
		Sheet sheet = wb.getSheet("Data");
		sheet.createRow(0).createCell(0).setCellValue("Roll No");
		sheet.getRow(0).createCell(1).setCellValue("Name");
		sheet.createRow(1).createCell(0).setCellValue(123456);
		sheet.getRow(1).createCell(1).setCellValue("Akshay Nilesh Khatavkar");
		
		FileOutputStream fos = new FileOutputStream("Hello123.xlsx");
		wb.write(fos);
		fos.close();
		wb.close();
	}
	
	private static String getCellVal(Row row,int col){
		try{
			return row.getCell(col).getStringCellValue().trim();
		}
		catch (IllegalStateException e){
			Double n = row.getCell(col).getNumericCellValue();
			return Integer.toString(n.intValue());
		}
	}
}
