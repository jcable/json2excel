package wainteg.json2excel

import org.apache.camel.builder.RouteBuilder
import org.apache.camel.component.jackson.JacksonDataFormat
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.*

/**
 * A Camel Groovy DSL Router
 */
class MyRouteBuilder extends RouteBuilder {

    /**
     * Let's configure the Camel routing rules using Groovy code...
     */
    void configure() {
    		JacksonDataFormat format = new JacksonDataFormat(Map.class);
    
        from("file:data")
        .to('log:start')
        .unmarshal(format)
        .process{
        	Workbook wb = new XSSFWorkbook()
		Font headingFont = wb.createFont();
		headingFont.setBold(true);
		XSSFCellStyle headingStyle = wb.createCellStyle();
		headingStyle.setFont(headingFont);
		it.in.body.each{ k,v -> 
			def rows = v
			if(k == 'broadcasts') rows = rows.broadcasts[0]
			XSSFSheet sheet = (XSSFSheet) wb.createSheet(k)
			def r = 0
			Row row = sheet.createRow(r++)
			row.setRowStyle(headingStyle);
			def c = 0
			def headings = rows.collectMany{it.keySet()}.toSet().toArray()
			headings.each{
				Cell cell = row.createCell(c++);
				cell.setCellValue(it)
			}
			rows.each{ p ->
				row = sheet.createRow(r++)
				c = 0
				headings.each{ h ->
					Cell cell = row.createCell(c++);
					if(p.containsKey(h)) {
						cell.setCellValue(p[h])	        				
					} else {
						cell.setCellValue(' ')	        				
					}
				}
			}	        		
			while(--c >= 0)
			    sheet.autoSizeColumn(c);	        		
		}
		OutputStream fileOut = new FileOutputStream("workbook.xlsx")
		wb.write(fileOut)
        }        
        .to('log:end')
    }
}
