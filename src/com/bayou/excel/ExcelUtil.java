package com.bayou.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelUtil {
	private static final String HSSFCell = null;
	public static void main(String[] args) throws Exception {
//		demo1();
		demo2();
	}
	/**
	 * 使用XSSFWordbook读取 xxx.xlsx文件
	 * @throws IOException
	 */
	public static void demo1() throws IOException {
		XSSFWorkbook xssfWord = new XSSFWorkbook("excelFIle/text2.xlsx");
		XSSFSheet sheet = xssfWord.getSheetAt(0);
		XSSFRow row = null;
		for(int i = 2; sheet.getRow(i) != null; i ++) {
			row = sheet.getRow(i);
			XSSFCell cell1 = row.getCell(0);
			XSSFCell cell2 = row.getCell(1);
			XSSFCell cell3 = row.getCell(2);
			XSSFCell cell4 = row.getCell(3);
			XSSFCell cell5 = row.getCell(4);
			XSSFCell cell6 = row.getCell(5);
			
			//从当前行到最后一行的数据
			System.out.println(cell2.toString());
			//获取单元格坐标比如 E3 A1
			String index = cell5.getReference();
			
			String rawValue1 = cell1.getStringCellValue();
			String rawValue2 = cell2.getStringCellValue();
			String rawValue3 = cell3.getStringCellValue();
			String rawValue4 = cell4.getStringCellValue();
			String rawValue5 = cell5.getStringCellValue();
			//读取数字单元格
			String rawValue6 = cell6.getRawValue();
			System.out.println(rawValue1 + "==" + rawValue2 + "==" + rawValue3 + "==" + rawValue4 + "==" + rawValue5 + "==" + rawValue6);
		}
	}
	/**
	 * 使用HSSFWorkbook读取xxx.xls文件
	 * @throws IOException
	 */
	@SuppressWarnings("deprecation")
	public static void demo2() throws IOException {
		FileInputStream fis = new FileInputStream(new File("excelFIle/text1.xls"));
		HSSFWorkbook hssfWorkbook = new HSSFWorkbook(fis);
//		HSSFSheet sheet = hssfWorkbook.getSheet("作业题库");
		HSSFSheet sheet = hssfWorkbook.getSheetAt(0);
		HSSFRow row = null;
		for(int i = 2; sheet.getRow(i) != null; i ++) {
			row = sheet.getRow(i);
			HSSFCell cell1 = row.getCell(0);
			HSSFCell cell2 = row.getCell(1);
			//读取数字的方法
//			int numericCellValue = (int) cell.getNumericCellValue();
			String value1 = cell1.getStringCellValue();
			String value2 = cell2.getStringCellValue();
			
			System.out.println(value1 + "==" + value2);
		}
		
		
		
	}
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
}
