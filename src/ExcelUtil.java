package hbu.poi;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.Map.Entry;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.Region;

public class ExcelUtil {
	
	public static HSSFWorkbook workbook=null;
	public static HSSFSheet sheet=null;
	public static HSSFRow row=null;
	public static HSSFCell cell=null;
	
	
	public static void createExcel(String templateName,String createName,int sheetNum,Map<String,String> mapParam,List<CellModel> cells){
		
		FileInputStream inputStream=null;
		FileOutputStream outputStream=null;

		//打开文件
		if(null!=templateName&&!"".equals(templateName)){
			File file=new File(templateName);
			if(file.exists()){
				try {
					inputStream=new FileInputStream(templateName);
					ByteArrayInputStream in=new ByteArrayInputStream(readInputStream(inputStream));
					workbook=new HSSFWorkbook(in);
					sheet=workbook.getSheetAt(sheetNum);
				} catch (FileNotFoundException e) {
					e.printStackTrace();
				} catch (IOException e) {
					System.out.println("打开文件失败");
					e.printStackTrace();
				}
			}
		}
		if(null!=createName&&!"".equals(createName)){
			try {
				outputStream=new FileOutputStream(createName);
				if(null==workbook){
					workbook=new HSSFWorkbook();
					sheet=workbook.createSheet();
				}
			} catch (FileNotFoundException e) {
				System.out.println("文件不存在");
				e.printStackTrace();
			}
		}
		if(null==workbook){
			return;
		}
		//写入map
		if(null!=mapParam){
			Map<String,String> params=getCellParams(sheet);
			Set<Entry<String,String>> paramSet=params.entrySet();
			for(Entry<String,String> entry:paramSet){
				String key=entry.getKey();
				String value=entry.getValue();
				String[] pos=value.split("#");
				cell=sheet.getRow(Short.parseShort(pos[0])).getCell(Short.parseShort(pos[1]));
				cell.setCellValue(mapParam.get(key));
			}
		}
		//写入cells
		if(null!=cells&&cells.size()>0){
			CellModel cellModel;
			for(int i=0;i<cells.size();i++){
				cellModel=cells.get(i);
				sheet.addMergedRegion(new Region(cellModel.getStartY(),
						(short) cellModel.getStartX(), cellModel.getEndY(),
						(short) cellModel.getEndX()));
				row=sheet.getRow(cellModel.getStartY());
				if(null==row){
					row=sheet.createRow(cellModel.getStartY());
				}
				cell=row.getCell(cellModel.getStartX());
				if(null==cell){
					cell=row.createCell(cellModel.getStartX());
				}
				HSSFCellStyle cellStyle=workbook.createCellStyle();
				cellStyle.cloneStyleFrom(cellModel.getCellStyle().getHssfCellStyle());
				cell.setCellStyle(cellStyle);
				cell.setCellValue(cellModel.getText());
			}
		}
		
		//输出文件
		try {
			workbook.write(outputStream);
			outputStream.close();
			inputStream.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
		
	}
	
	public static Map<String,String> getCellParams(HSSFSheet sheet){
		Map<String,String> params=new HashMap<String, String>();
		
		int firstRow=sheet.getFirstRowNum();
		int lastRow=sheet.getLastRowNum();
		
		HSSFRow row;
		HSSFCell cell;
		
		Pattern p=Pattern.compile("\\$\\{(.*)\\}");
		Matcher matcher;
		for(int rowIndex=firstRow;rowIndex<=lastRow;rowIndex++){
			row=sheet.getRow(rowIndex);
			if(null==row){
				continue;
			}
			int firstCell=row.getFirstCellNum();
			int lastCell=row.getLastCellNum();
			
			for(int cellIndex=firstCell;cellIndex<lastCell;cellIndex++){
				cell=row.getCell(cellIndex);
				if(null==cell){
					continue;
				}
				String  text=cell.getStringCellValue();
				if(null!=text&&!"".equals(text)){
					text=text.trim();
					matcher=p.matcher(text);
					String variable="";
					while(matcher.find()){
						variable=matcher.group(1);
					}
					if(!"".equals(variable)){
						params.put(variable,rowIndex+"#"+cellIndex);
					}
				}
			}
		}
		return params;
	}
	public static byte[] readInputStream(InputStream inputStream){
		ByteArrayOutputStream byteOs=new ByteArrayOutputStream();
		byte[] buffer=new byte[512];
		int len=0;
		try {
			while((len=inputStream.read(buffer))>0){
				byteOs.write(buffer,0,len);
			}
			byteOs.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
		return byteOs.toByteArray();
	}

	
}
