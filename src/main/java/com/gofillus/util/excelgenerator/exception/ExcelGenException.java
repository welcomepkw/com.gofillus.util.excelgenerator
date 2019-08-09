package com.gofillus.util.excelgenerator.exception;

public class ExcelGenException extends Exception{

	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;

	public ExcelGenException() {
		super();
	}

	public ExcelGenException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
		super(message, cause, enableSuppression, writableStackTrace);
	}

	public ExcelGenException(String message, Throwable cause) {
		super(message, cause);
	}

	public ExcelGenException(String message) {
		super(message);
	}

	public ExcelGenException(Throwable cause) {
		super(cause);
	}

	
}
