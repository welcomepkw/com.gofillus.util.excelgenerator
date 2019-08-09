package com.gofillus.util.excelgenerator.vo;

import java.util.List;

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
public class SheetVo {

	private String sheetName;
	private RowVo titles;
	private List<RowVo> datas; 
	
}
