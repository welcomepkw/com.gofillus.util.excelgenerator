package com.gofillus.util.excelgenerator.controller;

import java.io.IOException;

import javax.servlet.http.HttpServletResponse;

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
import com.gofillus.util.excelgenerator.vo.ExcelVo;

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
}
