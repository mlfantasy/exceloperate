package com.servlet;

import java.io.IOException;
import java.io.OutputStream;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellRangeAddress;

public class downloadExcel extends HttpServlet {

	
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;

	public downloadExcel() {
		super();
	}
	
	public void destroy() {
		super.destroy(); 		
	}
	
	public void doGet(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {
		//创建excel文档对象
		HSSFWorkbook wb = new HSSFWorkbook();
		
		//建立sheet对象
		HSSFSheet sheet = wb.createSheet("成绩表");
		
		//创建第一行
		HSSFRow row1 = sheet.createRow(0);
		
		//创建单元格
		HSSFCell cell1 = row1.createCell(0);
		
		//设置单元格的值
		cell1.setCellValue("学员成绩考试一览表");
		
		//设置合并行，四个参数分别代表起始行，终止行，起始列，截止列
		sheet.addMergedRegion(new CellRangeAddress(0,0,0,3));
		
		//创建第二行
        HSSFRow row2 = sheet.createRow(1);
		
		//创建单元格
	    row2.createCell(0).setCellValue("姓名");
	    row2.createCell(1).setCellValue("班级");;
	    row2.createCell(2).setCellValue("笔试成绩");;
	    row2.createCell(3).setCellValue("机试成绩");;
        
	  //创建第二行
        HSSFRow row3 = sheet.createRow(2);
		
		//创建单元格
	    row3.createCell(0).setCellValue("tony");
	    row3.createCell(1).setCellValue("a11401");
	    row3.createCell(2).setCellValue(52);
	    row3.createCell(3).setCellValue(67);
	    
	    OutputStream os = response.getOutputStream();
		response.reset();
		response.setHeader("Content-disposition", "attachment; filename=details.xls");
	    response.setContentType("application/msexcel");        
	    wb.write(os);
	    os.close();		
		
	}
	
	public void doPost(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {
		this.doGet(request, response);
	}
	
	public void init() throws ServletException {
	
	}

}
