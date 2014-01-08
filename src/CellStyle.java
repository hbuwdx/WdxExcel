package hbu.poi;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class CellStyle {
	
	private HSSFCellStyle hssfCellStyle;

	public HSSFCellStyle getHssfCellStyle() {
		if(null==hssfCellStyle){
			if(null==ExcelUtil.workbook){
				ExcelUtil.workbook=new HSSFWorkbook();
			}
			hssfCellStyle=ExcelUtil.workbook.createCellStyle();
		}
		return hssfCellStyle;
	}

	public void setHssfCellStyle(HSSFCellStyle hssfCellStyle) {
		this.hssfCellStyle = hssfCellStyle;
	}
}
