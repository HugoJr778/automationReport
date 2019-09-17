package enumService;

import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public enum StyleFormatFont {
	
	CENARIO {
		@Override
		public XSSFFont dataStyleWBFONT(XSSFWorkbook wb) {
			inicialize(wb);
			font.setColor(IndexedColors.INDIGO.getIndex());
			return font;
		}
	},
	
	STATUS_FONT {
		@Override
		public XSSFFont dataStyleWBFONT(XSSFWorkbook wb) {
			inicialize(wb);
			font.setColor(IndexedColors.RED.getIndex());
			return font;
		}
	}, 
	
	STATUS_DEV_COLOR {
		@Override
		public XSSFFont dataStyleWBFONT(XSSFWorkbook wb) {
			inicialize(wb);
			font.setColor(IndexedColors.DARK_GREEN.getIndex());
			return font;
		}
	};
	
	protected XSSFFont font;
	
	protected void inicialize(XSSFWorkbook wb) {
		this.font = wb.createFont();
	}
	
	public abstract XSSFFont dataStyleWBFONT(XSSFWorkbook wb);
}