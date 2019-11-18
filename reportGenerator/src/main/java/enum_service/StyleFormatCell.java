package enum_service;

import java.awt.Color;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public enum StyleFormatCell {
	
	BORDER {
		@Override
		public CellStyle dataStyleWBCELL(XSSFWorkbook wb) {
			inicialize(wb);
			style.setBorderBottom(BorderStyle.THIN);
			style.setBorderTop(BorderStyle.THIN);
			style.setBorderLeft(BorderStyle.THIN);
			style.setBorderRight(BorderStyle.THIN);
			return style;
		}
	},
	
	CENTER {
		@Override
		public CellStyle dataStyleWBCELL(XSSFWorkbook wb) {
			inicialize(wb);
			style.setAlignment(HorizontalAlignment.CENTER);
			style.setBorderBottom(BorderStyle.THIN);
			style.setBorderTop(BorderStyle.THIN);
			style.setBorderLeft(BorderStyle.THIN);
			style.setBorderRight(BorderStyle.THIN);
			return style;
		}
	},
	
	STATUS_COLOR_NOK {
		@Override
		public CellStyle dataStyleWBCELL(XSSFWorkbook wb) {
			inicialize(wb);
			style.setFillBackgroundColor(new XSSFColor(Color.RED));
			return style;
		}
	},
	
	STATUS_COLOR_OK {
		@Override
		public CellStyle dataStyleWBCELL(XSSFWorkbook wb) {
			inicialize(wb);
			style.setFillBackgroundColor(new XSSFColor(Color.GREEN));
			return style;
		}
	},

	STATUS_DEV {
		@Override
		public CellStyle dataStyleWBCELL(XSSFWorkbook wb) {
			inicialize(wb);
			style.setFillBackgroundColor(new XSSFColor(Color.LIGHT_GRAY));
			return style;
		}
	};
		
	protected XSSFCellStyle style;
	
	protected void inicialize(XSSFWorkbook wb) {
		this.style = wb.createCellStyle();
	}
	
	public abstract CellStyle dataStyleWBCELL(XSSFWorkbook wb);
}