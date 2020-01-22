/**
 * @author v.kwpark
 * @date 19.08.01
 * @desc ExcelGenerator.java 에서 엑셀로 생성하기 위한 프로토콜
 * @dependency jQuery >= 1.2
 */
var ExcelGenerator = new function(){
	
	this.CELL_TYPE = {
			_NONE : '_NONE'
			, NUMERIC : 'NUMERIC'
			, STRING : 'STRING'
			, FORMULA : 'FORMULA'
			, BLANK : 'BLANK'
			, BOOLEAN : 'BOOLEAN'
			, ERROR : 'ERROR'
	}
	
	/**
	 * @desc required thead and tbody
	 */
	this.convertTableToJson = function(arg, sheetName){
		// return new Promise(function(resolve, reject){
			var result = {titles:{columns:[]}, datas:[], sheetName:null};
			var thead = $(arg).find("thead");
			var tbody = $(arg).find("tbody");
			
			if(!thead || thead.lengh <= 0){
				// reject(new Error('empty thead'));
				return new Error('empty thead');
			}
			
			if(!sheetName){
				// reject(new Error('empty sheetName'));
				return new Error('empty sheetName');
			}
			
			// set sheet name
			result.sheetName = sheetName;
			
			// set title row
			var ths = $(arg).find("thead tr th");
			$(ths).each(function(index, obj){
				var titleObj = {};
				titleObj['backgroundColor'] = $(obj).css("background-color");
				titleObj['data'] = $(obj).text();
				titleObj['cellType'] = _getCellType($(obj).text());
				_getCellType(titleObj, $(obj).text());
				result.titles.columns.push(titleObj);
			});
			
			// set data rows
			var dataRows = $(arg).find("tbody tr");
			$(dataRows).each(function(index1, obj1){
				var tds = $(obj1).find("td");
				var dataRow = {columns:[]};
				$(tds).each(function(index2, obj2){
					var dataObj = {};
					dataObj['backgroundColor'] = $(obj2).css("background-color");
					if($(obj2).find("input[type='checkbox']").length > 0){
						dataObj['data'] = $(obj2).find("input[type='checkbox']").prop("checked")?'Y':'N';
					}else if($(obj2).find("input[type='radio']").length > 0){
						dataObj['data'] = $(obj2).find("input[type='radio']").prop("checked")?'Y':'N';
					}else if($(obj2).find("input[type='text']").length > 0){
						dataObj['data'] = $(obj2).find("input[type='text']").val();
					}else if($(obj2).find("input[type='number']").length > 0){
						dataObj['data'] = $(obj2).find("input[type='number']").val();
					}else{
						dataObj['data'] = $(obj2).text();
					}
					dataObj['cellType'] = _getCellType(dataObj['data']);
					dataRow.columns.push(dataObj);
				})
				
				result.datas.push(dataRow);
			});
			
			// resolve(result);
			return result;
		// });

	};
	
	this.combineSheetsToRequestData = function(arrayObj){
		return {sheets:arrayObj};
	};
	
	this.combineSheetsToRequestData = function(arrayObj, fileName){
		return {sheets:arrayObj, fileName:fileName};
	};
	
	_getCellType = function(data){
		
		if(!data){
			return ExcelGenerator.CELL_TYPE.BLANK;
		}
		
		if(Number(data)){
			if(data.indexOf('0') == 0){
				return ExcelGenerator.CELL_TYPE.STRING;
			}
			return ExcelGenerator.CELL_TYPE.NUMERIC;
		}
		
		return ExcelGenerator.CELL_TYPE.STRING;
	}
}