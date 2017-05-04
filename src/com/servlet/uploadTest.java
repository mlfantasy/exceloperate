package com.servlet;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Row;

public class uploadTest {
	
	public static void main(String[] args) {
		
		FileInputStream fis = null;
		
		try {
		    fis = new FileInputStream("d:\\details.xls");		
			Workbook wb = new HSSFWorkbook(fis);
			
			Sheet sheet = wb.getSheetAt(0);
			
			for(Row r : sheet){
				if(r.getRowNum()<1){
					continue;
				}
				System.out.println(r.getCell(0).toString());
				System.out.println(r.getCell(1).toString());
				System.out.println(r.getCell(2).toString());
				System.out.println(r.getCell(3).toString());
			}
		} catch (FileNotFoundException e) {			
			e.printStackTrace();
		} catch (IOException e) {			
			e.printStackTrace();
		} finally{
			try {
				fis.close();
			} catch (IOException e) {				
				e.printStackTrace();
			}
		}
		
		
		
		
	}

}
