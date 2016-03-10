package test;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

import org.junit.Before;
import org.junit.Test;

import test.ExcelUploadUtil;

public class ExcelUploadTest {

	
	public final String filePath ="파일경로";   
	public final String fileName = "파일명";   
	public Map<String, String> map  = new HashMap<String, String>();
	ExcelUploadUtil util = new ExcelUploadUtil();
	
	@Before
	public void init(){
		map.put("filePath", filePath);
		map.put("fileName", fileName);
		
	}
	@Test
	public void test_xlsx()throws Exception {
		ArrayList<ArrayList<Map<String,Object>>> list = util.excelImportXls_common(map);
		
		
		int i = 0 ;
		ArrayList<ArrayList<ArrayList<Object>>> lastArr = new ArrayList<ArrayList<ArrayList<Object>>>();
		for(ArrayList<Map<String,Object>> list2 : list ){
			ArrayList<ArrayList<Object>> arr =new ArrayList<ArrayList<Object>>();
			for(Map<String,Object> m: list2){
				ArrayList<Object> inArr = new ArrayList<Object>();
				for(int aa = 0; aa < m.size() ; aa++){
					inArr.add(m.get("no"+aa));
					System.out.println("key:no"+aa+",value:"+m.get("no"+aa));
				}
				arr.add(inArr);
			}
			lastArr.add(arr);
			i++;
		}
		
		
		
		
		
	}
	

}




