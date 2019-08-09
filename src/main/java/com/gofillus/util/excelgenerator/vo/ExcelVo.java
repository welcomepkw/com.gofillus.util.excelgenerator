package com.gofillus.util.excelgenerator.vo;

import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;


/**
 * 
 * @author v.kwpark
 * @date 19.08.01
 *
 */
@NoArgsConstructor
@AllArgsConstructor
@Builder
@Getter
@Setter
public class ExcelVo {

	@Builder.Default
	private Class<?> workbookType = XSSFWorkbook.class;
	private String fileName;
	private List<SheetVo> sheets;
}
