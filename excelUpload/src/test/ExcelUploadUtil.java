package test;

import java.io.File;
import java.io.FileInputStream;
import java.text.NumberFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelUploadUtil {
	
	
	
	public static final int BUFFER_SIZE = 8192;
	
	
	
	
	@SuppressWarnings({ "resource", "unused" })
	public ArrayList<ArrayList<Map<String,Object>>> excelImportXlsx_common(Map<String, String> fileMap) throws Exception{
		ArrayList<ArrayList<Map<String,Object>>> param_map = new ArrayList<ArrayList<Map<String,Object>>>();
		ArrayList<Map<String, Object>> paramList = null;
		NumberFormat nf = NumberFormat.getInstance();
		nf.setMinimumFractionDigits(0);//소수점 아래 최소 자리수
		nf.setMaximumFractionDigits(0);//소수점 아래 최대자리수
		
		XSSFWorkbook workbook = null;
		XSSFSheet sheet = null;
		XSSFRow row = null;
		XSSFCell cell = null;
        String filePath    =fileMap.get("filePath");
        String fileName  =fileMap.get("fileName");
        System.out.println(filePath + fileName);
	   FileInputStream excelFile = new FileInputStream(filePath + fileName);
	   File isExcelFile = new File(filePath + fileName);
	   
	   //파일이 존재하지 않을경우 EXCEPTION
	   if(!isExcelFile.isFile())throw new Exception("notFile");
        	   
        	   
       workbook  =  new XSSFWorkbook(excelFile);
       FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator(); // 수식 계산 시 필요!!!!!!
       //workBook 이 존재하지 않을경우 EXCEPTION 	   
       
	   if( workbook == null)throw new Exception("workbookIsNull");
	   
	   for(int sheetNum=0; sheetNum < workbook.getNumberOfSheets(); sheetNum++){
	    sheet = workbook.getSheetAt(sheetNum);	//멀티 시트에 순서대로 시트에 입힌다.
	    if(sheet ==null )continue; 						//시트가 존재하지 않을경우 continue
	    int nRowStartIndex = 0;			 //기록물철의 경우 실제 데이터가 시작되는 Row지정
	    String sheetNm = sheet.getSheetName();
        int nRowEndIndex   = sheet.getLastRowNum();       //기록물철의 경우 실제 데이터가 끝 Row지정
        int nColumnStartIndex = 0;						         //기록물철의 경우 실제 데이터가 시작되는 Column지정
        String szValue = "";										//시트에 값이 들어갈 부분.
               
        paramList = new ArrayList<Map<String,Object>>();
        
	       for( int i = nRowStartIndex; i <= nRowEndIndex ; i++){
			   int nColumnEndIndex = 0;			//기록물철의 경우 실제 데이터가 끝나는 Column지정
			   
			   if(nColumnEndIndex == sheet.getRow(0).getLastCellNum()){
				   nColumnEndIndex = sheet.getRow(i).getLastCellNum();
			   }else{
				   nColumnEndIndex = sheet.getRow(0).getLastCellNum();
			   }
			   row  = sheet.getRow(i);
			   String[] temp = new String[nColumnEndIndex];
	            	
	        	   for( int nColumn = nColumnStartIndex; nColumn < nColumnEndIndex ; nColumn++){
	        		   cell = row.getCell((short) nColumn);
	        		   if( cell == null){
	        			   temp[nColumn] = "";
	                       continue;
	                   }
	        		   switch(cell.getCellType()){
	            		   case XSSFCell.CELL_TYPE_NUMERIC : //0  number type
	            			   szValue = String.valueOf(nf.format(Math.round(cell.getNumericCellValue())));
	            			   break;
	            		   case XSSFCell.CELL_TYPE_STRING : //1   string type
	            			   szValue = cell.getStringCellValue();
	            			   break;
	            		   case XSSFCell.CELL_TYPE_FORMULA : //2
	            			   if(!"".equals(cell.toString())){
	            				   if(evaluator.evaluateFormulaCell(cell)==0){
	            			            double fddata = cell.getNumericCellValue();         
	            			            szValue = String.valueOf(nf.format(fddata));
	        			           }else if(evaluator.evaluateFormulaCell(cell)==1){
	        			        	   szValue = cell.getStringCellValue();
	        			           }else if(evaluator.evaluateFormulaCell(cell)==4){
	            			            boolean fbdata = cell.getBooleanCellValue();         
	            			            szValue = String.valueOf(fbdata);         
	        			           }
	            			   }
	            			   break;
	            		   case XSSFCell.CELL_TYPE_BOOLEAN: //4
	                             cell.getBooleanCellValue();  
	                             break;
	            		   default :
	            			   szValue = "";
	        		   }
	        		   temp[nColumn] = szValue;
	               }
        	   
	           Map<String,Object> m = new HashMap<String,Object>(); 
        	   //갯수가져와서 컬럼 만들어준다.
        	   for(int j=0;j<temp.length;j++)m.put("no"+j, "".equals(temp[j].trim()) ? sheetNm : temp[j]);
        	   
    		   paramList.add(m);
        	   
    	   }
	       param_map.add(paramList);
	   }
		return param_map;
	}
	
	@SuppressWarnings({ "unused", "deprecation" })
	public ArrayList<ArrayList<Map<String,Object>>> excelImportXls_common(Map<String, String> fileMap) throws Exception{

		ArrayList<ArrayList<Map<String,Object>>> param_map = new ArrayList<ArrayList<Map<String,Object>>>();
		ArrayList<Map<String, Object>> paramList = null;
		NumberFormat nf = NumberFormat.getInstance();
		nf.setMinimumFractionDigits(0);//소수점 아래 최소 자리수
		nf.setMaximumFractionDigits(0);//소수점 아래 최대자리수
		
		HSSFWorkbook workBook = null; // xls 버전
		HSSFSheet sheet = null; 
		HSSFRow row = null;
		HSSFCell cell = null; 
		
		String filePath    =fileMap.get("filePath");
		String fileName  =fileMap.get("fileName");
		
	    File isExcelFile = new File(filePath + fileName);
	    if(!isExcelFile.isFile())throw new Exception("notFile");//파일이 존재하지 않을경우 EXCEPTION
		   
		FileInputStream excelFile = new FileInputStream(filePath + fileName);  
		workBook =  new HSSFWorkbook(excelFile); // xls 버전
		FormulaEvaluator evaluator = workBook.getCreationHelper().createFormulaEvaluator(); // 수식 계산 시 필요!!!!!!
			
			
		if(workBook == null)throw new Exception("workbookIsNull");//workBook 이 존재하지 않을경우 EXCEPTION
		
		for(int sheetNum=0; sheetNum < workBook.getNumberOfSheets(); sheetNum++){
			String szValue = "";
			sheet = workBook.getSheetAt(sheetNum);
			
		    String sheetNm = sheet.getSheetName();
			paramList = new ArrayList<Map<String,Object>>();
			
			int rows = sheet.getPhysicalNumberOfRows();
			for (int r = 0; r <= rows; r++) {// 생성된 시트를 이용하여 그 행의 수만큼 돌면서 행을 하나씩 생성합니다.
				row = sheet.getRow(r);
				if(row != null){
					int cells = row.getLastCellNum();
					if(cells == sheet.getRow(0).getLastCellNum()){
						cells = row.getLastCellNum();
					}else{
						cells = sheet.getRow(0).getLastCellNum();
					}
					String[] temp = new String[cells];
			
					for (short c = 0; c < cells; c++) { // short 형입니다. 255개가 max!
						cell = row.getCell(c);
						if(cell != null){
							switch (cell.getCellType()) {	// 셀타입에 따라 출력 메소드 다름.
								case 0:
									 szValue = String.valueOf(nf.format(Math.round(cell.getNumericCellValue())));
									break;
								case 1:
									szValue = cell.getStringCellValue();
									break;
								case Cell.CELL_TYPE_FORMULA :
									 if(!"".equals(cell.toString())){
											 if(evaluator != null){
				                				   if(evaluator.evaluateFormulaCell(cell)==0){
				                			            double fddata = cell.getNumericCellValue();         
				                			            szValue = String.valueOf(nf.format(fddata));
			                			           }else if(evaluator.evaluateFormulaCell(cell)==1){
			                			        	   szValue = cell.getStringCellValue();
			                			           }else if(evaluator.evaluateFormulaCell(cell)==4){
				                			            boolean fbdata = cell.getBooleanCellValue();         
				                			            szValue = String.valueOf(fbdata);         
			                			           }
			                			   }else{
			                				   szValue = "";
			                			   }
									 }
	                			   break;
								default :
	                			   szValue = "";
							}
						}else{ 
							szValue = "";
						}
						temp[c] = szValue;
					}
                	   
					 Map<String,Object> m = new HashMap<String,Object>(); 
		        	   //갯수가져와서 컬럼 만들어준다.
					for(int j=0;j<temp.length;j++)m.put("no"+j, "".equals(temp[j].trim()) ? sheetNm : temp[j]);
					
            	   paramList.add(m);
				}
				
			}
			 param_map.add(paramList);
		}
		
		return param_map;
	}
	
	
	
	
	
}
