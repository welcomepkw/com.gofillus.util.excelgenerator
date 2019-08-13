package com.gofillus.util.excelgenerator.controller;

import java.io.IOException;
import java.util.Arrays;

import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;

import com.fasterxml.jackson.core.JsonParseException;
import com.fasterxml.jackson.databind.JsonMappingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.gofillus.util.excelgenerator.core.ExcelGenerator;
import com.gofillus.util.excelgenerator.exception.ExcelGenException;
import com.gofillus.util.excelgenerator.vo.ColumnVo;
import com.gofillus.util.excelgenerator.vo.ExcelVo;
import com.gofillus.util.excelgenerator.vo.RowVo;
import com.gofillus.util.excelgenerator.vo.SheetVo;

@Controller
@RequestMapping("/exam")
public class ExampleController {

	@Autowired
	private ExcelGenerator excelGen;
	
	@RequestMapping(value = {"/genJsonToExcel"}, method = {RequestMethod.POST})
	public void genJsonToExcel(
			@RequestParam(name = "excelData", required = true) String excelData
			, HttpServletResponse response
			) throws JsonParseException, JsonMappingException, IOException, ExcelGenException {
		
		ObjectMapper mapper = new ObjectMapper();
		ExcelVo excelVo = mapper.readValue(excelData, ExcelVo.class);
		
		excelGen.genExcelFile(excelVo, response);
	}
	
	@RequestMapping(value = {"/test"}, method = {RequestMethod.GET})
	public void test(
				HttpServletResponse response
			) throws ExcelGenException, IOException {
		
		String style = "rgb(100,100,100)";
		
		ColumnVo title1 = ColumnVo.builder()
				.data("title1")
				.horizontalAlignment(HorizontalAlignment.CENTER)
				.cellType(CellType.STRING)
				.backgroundColor(style)
				.build();
		ColumnVo title2 = ColumnVo.builder()
				.data("title2")
				.horizontalAlignment(HorizontalAlignment.CENTER)
				.cellType(CellType.STRING)
				.backgroundColor(style)
				.build();
		
		RowVo titleRow = RowVo.builder()
				.columns(Arrays.asList(title1, title2))
				.build();
		
		
		ColumnVo data1 = ColumnVo.builder()
				.data("data1")
				.horizontalAlignment(HorizontalAlignment.LEFT)
				.cellType(CellType.STRING)
				.build();
		ColumnVo data2 = ColumnVo.builder()
				.data("data2")
				.horizontalAlignment(HorizontalAlignment.LEFT)
				.cellType(CellType.STRING)
				.build();
		
		RowVo dataRow = RowVo.builder()
				.columns(Arrays.asList(data1, data2))
				.build();
				
		
		SheetVo sheet1 = SheetVo.builder()
				.sheetName("sheet1")
				.titles(titleRow)
				.datas(Arrays.asList(dataRow))
				.build();
		
		ExcelVo excelVo = ExcelVo.builder()
				.fileName("test.xls")
				.workbookType(HSSFWorkbook.class)
				.sheets(Arrays.asList(sheet1))
				.build();
		
		excelGen.genExcelFile(excelVo, response);
	}
}
