package com.excelReadWrite.test;

import org.testng.annotations.Test;

import Utils.ExcelUtils;

public class ExcelReadData {
	
	String fileName="Credentials";
	ExcelUtils excel;
	@Test
	public void readRowCount() {
		excel=new ExcelUtils(fileName);
		int rowCount=excel.getRowCount();		
		System.out.println("rowCount"+rowCount);
		
		
	}
	@Test
	public void readCellData() {
		excel=new ExcelUtils(fileName);
		 		
		System.out.println("User Name 1 : "+excel.getCellDataString(1,0));
		
		System.out.println("Pass Word 2 : "+excel.getCellDataString(2,1));
		
	}
}
