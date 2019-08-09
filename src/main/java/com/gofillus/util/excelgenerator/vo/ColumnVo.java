package com.gofillus.util.excelgenerator.vo;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

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
public class ColumnVo {

	private String data;
	private CellType cellType;
	private CellStyle cellStyle;
	
	// for json to excel
	private String backgroundColor;
	private HorizontalAlignment horizontalAlignment;
}
