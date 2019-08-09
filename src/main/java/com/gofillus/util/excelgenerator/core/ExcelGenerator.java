package com.gofillus.util.excelgenerator.core;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

import com.gofillus.util.excelgenerator.exception.ExcelGenException;
import com.gofillus.util.excelgenerator.vo.ColumnVo;
import com.gofillus.util.excelgenerator.vo.ExcelVo;
import com.gofillus.util.excelgenerator.vo.RowVo;
import com.gofillus.util.excelgenerator.vo.SheetVo;

/**
 * 
 * @author kay park
 * @date 19.08.02
 *
 */
@Component
public class ExcelGenerator {

	private Map<Short, CellStyle> definedCellStyle;  
	
	/**
	 * 
	 * @param arg
	 * @param response
	 * @throws ExcelGenException
	 * @throws IOException
	 */
	public void genExcelFile(ExcelVo arg, HttpServletResponse response) throws ExcelGenException, IOException {

		definedCellStyle = new HashMap<Short, CellStyle>();
		
		if (arg.getWorkbookType() == null) {
			throw new ExcelGenException("empty workbook type");
		}

		Workbook wb = null;

		if (arg.getWorkbookType().equals(HSSFWorkbook.class)) // xls
			wb = new HSSFWorkbook();

		if (arg.getWorkbookType().equals(XSSFWorkbook.class)) // xlsx
			wb = new XSSFWorkbook();

		for (SheetVo sheetData : arg.getSheets()) {

			if (sheetData.getSheetName() == null) {

			}

			Sheet sheet = wb.createSheet(sheetData.getSheetName());

			// add title row
			Row row = sheet.createRow(0);
			Cell cell = null;
			for (int i = 0; i < sheetData.getTitles().getColumns().size(); i++) {
				ColumnVo columnVo = sheetData.getTitles().getColumns().get(i);
				cell = row.createCell(i);
				setCellType(columnVo, cell, wb);
				setCellStyle(cell, wb, columnVo);
				setCellAlign(cell, columnVo);
				
				// set column width
				sheet.autoSizeColumn(i);
				sheet.setColumnWidth(i, sheet.getColumnWidth(i) + 512);
			}

			// add data rows
			if (sheetData.getDatas() != null) {
				for (int i = 0; i < sheetData.getDatas().size(); i++) {
					row = sheet.createRow(i + 1);
					RowVo rowVo = sheetData.getDatas().get(i);
					for (int j = 0; j < rowVo.getColumns().size(); j++) {
						cell = row.createCell(j);
						setCellType(rowVo.getColumns().get(j), cell, wb);
						
						setCellStyle(cell, wb, rowVo.getColumns().get(j));
						setCellAlign(cell, rowVo.getColumns().get(j));
					}

				}
			}

		}

		response.setContentType("ms-vnd/excel");
		response.setHeader("Content-Disposition", "attachment;filename=" + arg.getFileName());
		wb.write(response.getOutputStream());
		wb.close();
	}

	/**
	 * 
	 * @param arg
	 * @param cell
	 * @param wb
	 * @throws ExcelGenException
	 */
	private void setCellType(ColumnVo arg, Cell cell, Workbook wb) throws ExcelGenException {

		if (arg.getCellType() == null) {
			throw new ExcelGenException("empty cell type");
		}

		cell.setCellType(arg.getCellType());
		if (arg.getCellStyle() != null) {
			cell.setCellStyle(arg.getCellStyle());
		}
		switch (arg.getCellType()) {
		case STRING:
			cell.setCellValue(arg.getData());
			break;
		case NUMERIC:
			cell.setCellValue(Integer.parseInt((arg.getData())));
			break;
		case BOOLEAN:
			cell.setCellValue(Boolean.parseBoolean((arg.getData())));
			break;
		case FORMULA:
			cell.setCellValue(arg.getData());
			break;
		case BLANK:
			cell.setCellValue("");
			break;
		case ERROR:
			cell.setCellValue("");
			break;
		case _NONE:
			cell.setCellValue("");
			break;
		}

	}

	/**
	 * 		
	 * @param cell
	 * @param arg
	 */
	private void setCellAlign(Cell cell, ColumnVo arg) {
		if(arg.getHorizontalAlignment() == null) {
			return;
		}
		cell.getCellStyle().setAlignment(arg.getHorizontalAlignment());
	}
	
	/**
	 * 
	 * @param cell
	 * @param wb
	 * @param arg
	 * @throws ExcelGenException
	 */
	private void setCellStyle(Cell cell, Workbook wb, ColumnVo arg) throws ExcelGenException {
		
		if(arg.getCellStyle() != null) {
			cell.setCellStyle(arg.getCellStyle());
			return;
		}
		
		if(arg.getBackgroundColor() == null) {
			return;
		}
		
		// came from json 
		Pattern RGB = Pattern.compile("rgb\\(\\s*(\\d+)\\s*,\\s*(\\d+)\\s*,\\s*(\\d+)\\s*\\)");
		Pattern RGBA = Pattern.compile("rgba\\(\\s*(\\d+)\\s*,\\s*(\\d+)\\s*,\\s*(\\d+)\\s*,\\s*(\\d+)\\s*\\)");
		Matcher matRGB = RGB.matcher(arg.getBackgroundColor());
		Matcher matRGBA = RGBA.matcher(arg.getBackgroundColor()); 
		if (!matRGB.find() && !matRGBA.find()) { 
			throw new ExcelGenException("not matched rgb type");
		}
		
		if(matRGBA.find(0)) {
			matRGB = matRGBA;
			if(Integer.parseInt(matRGBA.group(4)) == 0)
				return;
		}
		matRGB.find(0);
		
		int r = Integer.parseInt(matRGB.group(1));
		int g = Integer.parseInt(matRGB.group(2));
		int b = Integer.parseInt(matRGB.group(3));
		byte[] rgb = new byte[] { (byte)r, (byte)g, (byte)b };		
		
		
		CellStyle cellcolorstyle = null;
		
		if (wb instanceof XSSFWorkbook) {
			
			
			XSSFColor customColor = new XSSFColor(rgb, new DefaultIndexedColorMap());
			
			cellcolorstyle = definedCellStyle.get(customColor.getIndex());
			
			if(cellcolorstyle == null) {
				cellcolorstyle = wb.createCellStyle();
				XSSFCellStyle xssfcellcolorstyle = (XSSFCellStyle) cellcolorstyle;
				xssfcellcolorstyle.setFillForegroundColor(customColor);
				xssfcellcolorstyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				definedCellStyle.put(customColor.getIndex(), xssfcellcolorstyle);
			}
			
		} else if (wb instanceof HSSFWorkbook) {
			
			HSSFWorkbook hssfworkbook = (HSSFWorkbook) wb;
			HSSFPalette palette = hssfworkbook.getCustomPalette();
			HSSFColor customColor = palette.findSimilarColor(rgb[0], rgb[1], rgb[2]);
			
			if(customColor == null) {
				throw new ExcelGenException("not exist similar color");
			}
			
			cellcolorstyle = definedCellStyle.get(customColor.getIndex());
			if(cellcolorstyle == null) {
				cellcolorstyle = wb.createCellStyle();
				cellcolorstyle.setFillForegroundColor(customColor.getIndex());
				cellcolorstyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				definedCellStyle.put(customColor.getIndex(), cellcolorstyle);
			}
			
			
		}

		
		if(cellcolorstyle.getFillForegroundColor() == HSSFColor.HSSFColorPredefined.WHITE.getIndex()) { return; }
		cell.setCellStyle(cellcolorstyle);
		
	}
	
}
