package hbu.poi;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;



public class Test {
	public static void main(String[] args) {
		String templateName="test.xls";
		String createName="test1.xls";
		
		List<CellModel> cells=new ArrayList<CellModel>();
		CellModel model;
		CellStyle cellStyle=new CellStyle();
		cellStyle.getHssfCellStyle().setAlignment(HSSFCellStyle.ALIGN_CENTER);
		cellStyle.getHssfCellStyle().setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		
		
		for(int i=0;i<1;i++){
			model=new CellModel();
			model.setStartX(i);
			model.setStartY(i);
			model.setEndX(i+3);
			model.setEndY(i+3);
			model.setCellStyle(cellStyle);
			model.setText("hello"+i);
			cells.add(model);
		}
		
		Map<String,String> mapParams=new HashMap<String, String>();
		mapParams.put("hello","你好");
		mapParams.put("hbu","河北大学");
		ExcelUtil.createExcel(templateName, createName, 0, mapParams, cells);
		
	}
}
