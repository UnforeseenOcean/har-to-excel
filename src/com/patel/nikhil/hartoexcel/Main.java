package com.patel.nikhil.hartoexcel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;

import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
import org.json.simple.parser.ParseException;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
 
@SuppressWarnings("unchecked")
public class Main {
	
	public static final String harFilePath = "/Users/nikhil/Desktop/phresco.har";
	
	public static final String requestFilePath = "/Users/nikhil/Desktop/request.xls";
	public static final String requestSheetName = "Request Headers";
	
	public static final String responseFilePath = "/Users/nikhil/Desktop/response.xls";
	public static final String responseSheetName = "Response Headers";
	
	public static void main(String[] args) {
		ArrayList<HashMap<String, HashMap<String,String>>> requestResult = new ArrayList<HashMap<String, HashMap<String, String>>>();
		ArrayList<HashMap<String, HashMap<String,String>>> responseResult = new ArrayList<HashMap<String, HashMap<String, String>>>();
		ArrayList<String> requestHeadersNameList = new ArrayList<String>();
		ArrayList<String> responseHeadersNameList = new ArrayList<String>();
		
		JSONParser parser = new JSONParser();
		try {
			Object obj = parser.parse(new FileReader(harFilePath));
			JSONObject jsonObject = (JSONObject) obj;
			JSONObject log = (JSONObject)jsonObject.get("log");
			JSONArray entries = (JSONArray)log.get("entries");
    		
			Iterator<JSONObject> entryIterator = entries.iterator();
			while(entryIterator.hasNext()) {
				JSONObject entry = entryIterator.next();
				
				/**
				 * Request
				 */
				JSONObject request = (JSONObject)entry.get("request");
				String requestUrl = (String)request.get("url");
				JSONArray requestHeaders = (JSONArray)request.get("headers");
				HashMap<String, String> requestHeadersMap = getHeadersMap(requestHeaders, requestHeadersNameList);

				HashMap<String, HashMap<String, String>> requestEntryMap = new HashMap<String, HashMap<String,String>>();
				requestEntryMap.put(requestUrl, requestHeadersMap);
				
				requestResult.add(requestEntryMap);
				
				/**
				 * Response
				 */
				JSONObject response = (JSONObject)entry.get("response");
				JSONArray responseHeaders = (JSONArray)response.get("headers");
				HashMap<String, String> responseHeadersMap = getHeadersMap(responseHeaders, responseHeadersNameList);

				HashMap<String, HashMap<String, String>> responseEntryMap = new HashMap<String, HashMap<String,String>>();
				responseEntryMap.put(requestUrl, responseHeadersMap);
				
				responseResult.add(responseEntryMap);
			}
			
			writeToExcel(requestSheetName, requestHeadersNameList, requestResult, requestFilePath);
			writeToExcel(responseSheetName, responseHeadersNameList, responseResult, responseFilePath);
			
		} catch (IOException | ParseException e) {
			e.printStackTrace();
		}
	}
	
	/**
	 * Creates a Map - headerName : headerValue and add unique Header names to ArrayList 	
	 * @param headers
	 * @param headersNameList
	 * @return headersMap
	 */
	private static HashMap<String, String> getHeadersMap(JSONArray headers, ArrayList<String> headersNameList) {
		HashMap<String, String> headersMap = new HashMap<String, String>();
		Iterator<JSONObject> requestHeaderIterator = headers.iterator();
		while(requestHeaderIterator.hasNext()) {
			JSONObject header = requestHeaderIterator.next();
			String headerName = (String)header.get("name");
			String headerValue = (String)header.get("value");
			headersMap.put(headerName, headerValue);
			if (!headersNameList.contains(headerName)) {
				headersNameList.add(headerName);
			}
		}
		return headersMap;
	}
	
	/**
	 * Write the data to Excel file
	 * @param sheetName
	 * @param headersNameList
	 * @param result
	 * @param filePath
	 * @throws IOException
	 */
	private static void writeToExcel(String sheetName, ArrayList<String> headersNameList, ArrayList<HashMap<String, HashMap<String,String>>> result, String filePath) throws IOException {
		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFSheet sheet = workbook.createSheet(sheetName);

		int rowNum = 0;
		Row headerRow = sheet.createRow(rowNum++);
		
		Cell urlHeaderCell = headerRow.createCell(0);
		urlHeaderCell.setCellValue("URL");
		
		int cellNum = 1;
		for (String header : headersNameList) {
			Cell headerCell = headerRow.createCell(cellNum++);
			headerCell.setCellValue(header);
		}
		
		for (HashMap<String, HashMap<String, String>> entry : result) {
			Row row = sheet.createRow(rowNum++);
			for (String url : entry.keySet()) {
				Cell urlCell = row.createCell(0);
				urlCell.setCellValue(url);
				cellNum = 1;
				HashMap<String, String> headersMap = entry.get(url);
				for (int i = 0; i < sheet.getRow(0).getLastCellNum(); i++) {
					if ( sheet.getRow(0).getCell(cellNum) == null) continue;
					String rowHeader = sheet.getRow(0).getCell(cellNum).getStringCellValue();
					Cell cell = row.createCell(cellNum++);
					cell.setCellValue(headersMap.get(rowHeader));
				}
			}
		}
		FileOutputStream out = new FileOutputStream(new File(filePath));
	    workbook.write(out);
	    out.close();
	}
}