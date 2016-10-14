import java.io.File;
import java.io.IOException;
import jxl.*;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class ExcelConnection {

	public static void main(String[] args) {
		
		int fileStart = 1, fileEnd = 69;
		
		try{
			String code,codeName,outCodeName;
			for(int i = fileStart; i<= fileEnd; i++)
			{
				if(i/10 < 1)
					code = "00000" + i;
				else if(i/10 < 10)
					code = "0000" + i;
				else code = "000" + i;
				codeName = "(" + code + ".SZ)";
				outCodeName = "[" + code + ".SZ]";
				File inFile = new File("input/-Unlicensed-"+"离任高管" + codeName + ".xls");
				
				if(inFile.exists())
				{
					File exlOutFile = new File("output/"+outCodeName + ".xls");
					WritableWorkbook writableWorkbook = Workbook.createWorkbook(exlOutFile);
					WritableSheet writableSheet = writableWorkbook.createSheet("Sheet1", 0);
					preWrite(writableSheet);
					int row = 0;
					
					Workbook book = Workbook.getWorkbook(inFile);
					Sheet sheet = book.getSheet(0);
					
					int j = 1;
					while(j < sheet.getRows()){
						String content = readCell(0,j,sheet);
						row++;
						if(j == 112)
							System.out.println(j);
						writeCell(0,row,writableSheet,code);
						writeCell(1,row,writableSheet,content);
						writeCell(2,row,writableSheet,"离任");
						for(int k = 1; k<=8; k++){
							content = readCell(k,j,sheet);
							writeCell(k+2,row,writableSheet,content);
						}
						j++;
					}
					inFile = new File("input/-Unlicensed-"+"现任管理层" + codeName + ".xls");
					book = Workbook.getWorkbook(inFile);
					sheet = book.getSheet(0);
					
					j = 1;
					
					while(j < sheet.getRows()){
						String content = readCell(0,j,sheet);
						String job = "";
						if("董事会".equals(content) == true){
							job = "董事会";
							j = j + 2;
							content = readCell(0,j,sheet);
						}
						else if ("监事会".equals(content) == true){
							job = "监事会";
							j = j + 2;
							content = readCell(0,j,sheet);
						}
						else if("高管".equals(content) == true){
							job = "高管";
							j = j + 2;
							content = readCell(0,j,sheet);
						}
						row++;
						writeCell(0,row,writableSheet,code);
						writeCell(1,row,writableSheet,content);
						writeCell(2,row,writableSheet,job);
						for(int k = 1; k<=2; k++){
							content = readCell(k,j,sheet);
							writeCell(k+2,row,writableSheet,content);
						}
						writeCell(5,row,writableSheet,"0");
						for(int k = 3; k<=7; k++){
							content = readCell(k,j,sheet);
							writeCell(k+3,row,writableSheet,content);
						}
						
						j++;
					
					}
					writableWorkbook.write();
			        writableWorkbook.close();
				}
			}
		
		}catch (IOException e) {
            e.printStackTrace();
		}catch (RowsExceededException e) {
            e.printStackTrace();
		}catch (WriteException e) {
            e.printStackTrace();
        }catch (Exception e){
        	e.printStackTrace();
        }
	}
	
	private static String readCell(int i, int j, Sheet sheet){
		
		Cell cell = sheet.getCell(i,j);
		String st = cell.getContents();
		return st;
	}
	
	private static void writeCell(int i, int j, WritableSheet sheet, String content) throws RowsExceededException, WriteException{
		
		Label label = new Label (i,j,content);
		sheet.addCell(label);
	}
	
	private static void preWrite(WritableSheet writableSheet) throws RowsExceededException, WriteException{
		
		 Label label = new Label(0, 0, "代码");
		 writableSheet.addCell(label);
		 label = new Label(1, 0, "姓名");
		 writableSheet.addCell(label);
		 label = new Label(2, 0, "属性");
		 writableSheet.addCell(label);
		 label = new Label(3, 0, "职务");
		 writableSheet.addCell(label);
		 label = new Label(4, 0, "任职日期");
		 writableSheet.addCell(label);
		 label = new Label(5, 0, "离职日期");
		 writableSheet.addCell(label);
		 label = new Label(6, 0, "性别");
		 writableSheet.addCell(label);
		 label = new Label(7, 0, "国籍");
		 writableSheet.addCell(label);
		 label = new Label(8, 0, "学历");
		 writableSheet.addCell(label);
		 label = new Label(9, 0, "出生年份");
		 writableSheet.addCell(label);
		 label = new Label(10, 0, "个人简历");
		 writableSheet.addCell(label);
	}
}
