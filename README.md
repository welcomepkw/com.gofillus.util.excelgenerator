# excel generator
This Library help to generate Excel file by Json data and Java Object.
Support xls, xlsx.

이 라이브러리는 Json data 와 Java Object 를 활용해서 Excel 파일을 생성합니다.

## Why make this
- Use agin later :)

## target JDK
- jdk 1.8

## Included dependencies
- poi-4.0.jar
- poi-ooxml-4.0.jar
- lombok
- Jquery 1.2 <=

## how to use
 1. Example 1. Use Java Object
     - Copy core, exception, vo package and java files from com.gofillus.util.excelgenerator to your project and change packages from java files.
     - Use like ExampleController.java -> test     
     
 2. Example 2. Use Json data
    - copy ExcelGenerator.js from /src/main/resources And include it to your html or jsp.
    - Use like Index.html -> excelGen()
