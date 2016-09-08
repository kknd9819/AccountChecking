package com.yidepiao.check;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.HashMap;
import java.util.Map;
import java.util.Scanner;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;  
  
public class ReadWriteExcel {  
  
    private static final String EXCEL_XLS = "xls";  
    private static final String EXCEL_XLSX = "xlsx";  
 

  
    /** 
     * 判断Excel的版本,获取Workbook 
     * @param in 
     * @param filename 
     * @return 
     * @throws IOException 
     */  
    public static Workbook getWorkbok(InputStream in,File file) throws IOException{  
        Workbook wb = null;  
        if(file.getName().endsWith(EXCEL_XLS)){  //Excel 2003  
            wb = new HSSFWorkbook(in);  
        }else if(file.getName().endsWith(EXCEL_XLSX)){  // Excel 2007/2010  
            wb = new XSSFWorkbook(in);  
        }  
        return wb;  
    }  
  
    /** 
     * 判断文件是否是excel 
     * @throws Exception  
     */  
    public static void checkExcelVaild(File file) throws Exception{  
        if(!file.exists()){  
            throw new Exception("文件不存在");  
        }  
        if(!(file.isFile() && (file.getName().endsWith(EXCEL_XLS) || file.getName().endsWith(EXCEL_XLSX)))){  
            throw new Exception("文件不是Excel");  
        }  
    }  
  
    /** 
     * 读取Excel测试，兼容 Excel 2003/2007/2010 
     * @throws Exception  
     */  
    public static void main(String[] args) throws Exception {  
    	int k = 0;
//    	Scanner sc = new Scanner(System.in);
//    	System.out.println("请输入第一份excel表格文件的路径");
//    	String file1 = sc.nextLine();
//    	System.out.println("请设置第一份excel表格跳过头部几行数据?  (输入一个数字)");
//    	int count1 = Integer.parseInt(sc.nextLine());
//    	System.out.println("请输入第二份excel表格文件的路径");
//    	String file2 = sc.nextLine();
//    	System.out.println("请设置第二份excel表格跳过头部几行数据?  (输入一个数字)");
//    	int count2 = Integer.parseInt(sc.nextLine());
    	
        Map<String,String>map1 = new HashMap<String,String>();
        Map<String,String>map2 = new HashMap<String,String>();
    	startRead("C:/Users/ZZ/Desktop/易得票8月份销售明细_东莞潇湘.xls", 3,map1,map2);
    	startRead1( "C:/Users/ZZ/Desktop/潇湘国际影城（东莞店）结算表_201608.xlsx", 1, map1, map2);
//    	System.out.println(map1.size());
//    	System.out.println(map2.size());
    	
//    	  System.out.println("amount1:" + amount1);
//    	  System.out.println("amount2:" + amount2);
    	
//    	Set<String> set = map1.keySet();
//  //  	int i=1;
//    	for (String key1 : set) {
//    		 if(!key1.equals("订单号")){
//       		  k +=	 Integer.parseInt(map1.get(key1));
//       		}
//    		String key2 = key1.substring(3);
//			if(map2.containsKey(key2)){
//				String value1 = map1.get(key1);
//				String value2 = map2.get(key2);
////				System.out.println(i + "-" +  key2 +"-"+ value1 + " - " + value2);
////				i++;
//				if(!value1.equals(value2)){
//					System.out.println("订单号为：" + key1 + "，价格1：" + map1.get(key1) + "，价格2：" + map2.get(key2));
//				}
//			}else{
//				System.out.println(key1);
//			}
//		}
//    	System.out.println(k);
//    	
    	
    	
    	
    	Set<String> set1 = map1.keySet();
		for (String key1 : set1) {
			String key2 = key1.substring(3);
			if (!map2.containsKey(key2)) {
				System.out.println(key1);
			}
		}
    	
    	
		Set<String> set2 = map2.keySet();
		for (String key2 : set2) {
			String key1 = "000" + key2;
			if (!map1.containsKey(key1)) {
				System.out.println(key2);
			}
		}
    	
    	
    	
    	
    	
    	
    	
    	
    	
  //  	sc.close();
    }
    
    
    public static  int amount1= 0;
    public static  int amount2= 0;
 
    public static void writeSql(String rowValue,Map map1,Map map2) throws IOException{  
    	
        String[] sqlValue = rowValue.split("#") ;
    
       	if(sqlValue.length >= 14 && !"已退".equals(sqlValue[14])) {
       		
       		if(!"票价".equals(sqlValue[12])){
       			amount1 += Float.parseFloat(sqlValue[12]);
       		}
       		
       		String ordernum = sqlValue[2];
       		if(map1.containsKey(ordernum)){
       			if(!map1.get(ordernum).equals(sqlValue[12])){
       				System.out.println("ordernum:" + ordernum);
       			}
       		}
       		
       		map1.put(sqlValue[2], sqlValue[12]);
        }
        
    //    sql="INSERT INTO table_name (列名1) VALUES("+ sqlValue[0].trim() + ");"+"\n";  

//        try {  
//            bw.write(sql);  
//            bw.newLine();  
//        } catch (IOException e) {  
//            e.printStackTrace();  
//        }  
    }  
    
    public static void startRead(String inpath,int beginRow,Map map1,Map map2)  throws Exception{
    	 SimpleDateFormat fmt = new SimpleDateFormat("yyyy-MM-dd");  
  //       BufferedWriter bw = new BufferedWriter(new FileWriter(new File(outpath)));  
         try {  
             // 同时支持Excel 2003、2007  
             File excelFile = new File(inpath); // 创建文件对象  
             FileInputStream is = new FileInputStream(excelFile); // 文件流  
             checkExcelVaild(excelFile);  
             Workbook workbook = getWorkbok(is,excelFile);  
             //Workbook workbook = WorkbookFactory.create(is); // 这种方式 Excel2003/2007/2010都是可以处理的  
   
             int sheetCount = workbook.getNumberOfSheets(); // Sheet的数量  
          //   System.out.println("Sheet的数量为  :" + sheetCount);
             /** 
              * 设置当前excel中sheet的下标：0开始 
              */  
             Sheet sheet = workbook.getSheetAt(0);   // 遍历第一个Sheet  
   
             // 为跳过第一行目录设置count  
             int count = beginRow;  
             
             for (Row row : sheet) {  
                 // 跳过第一行的目录  
                 if(count == 0){  
                     count++;  
                     continue;  
                 }  
                 // 如果当前行没有数据，跳出循环  
                 if(row.getCell(0) == null || row.getCell(0).toString().equals("")){  
                     return ;  
                 }  
                 String rowValue = "";  
                 
                 int col = 0;
                 
                 for (Cell cell : row) {  
                     if(cell.toString() == null){  
                         continue;  
                     }  
                     int cellType = cell.getCellType();  
                     String cellValue = "";  
                     switch (cellType) {  
                         case Cell.CELL_TYPE_STRING:     // 文本  
                             cellValue = cell.getRichStringCellValue().getString() + "#" ;                             
                             break;  
                         case Cell.CELL_TYPE_NUMERIC:    // 数字、日期  
                             if (DateUtil.isCellDateFormatted(cell)) {  
                                 cellValue = fmt.format(cell.getDateCellValue()) + "#" ;  
                             } else {  
                                 cell.setCellType(Cell.CELL_TYPE_STRING);  
                                 cellValue = String.valueOf(cell.getRichStringCellValue().getString()) + "#" ;  
                             }  
                             break;  
                         case Cell.CELL_TYPE_BOOLEAN:    // 布尔型  
                             cellValue = String.valueOf(cell.getBooleanCellValue()) + "#"  ;  
                             break;  
                         case Cell.CELL_TYPE_BLANK: // 空白  
                             cellValue = cell.getStringCellValue() + "#" ;  
                             break;  
                         case Cell.CELL_TYPE_ERROR: // 错误  
                             cellValue = "错误#" ;  
                             break;  
                         case Cell.CELL_TYPE_FORMULA:    // 公式  
                             // 得到对应单元格的公式  
                             //cellValue = cell.getCellFormula() + "#";  
                             // 得到对应单元格的字符串  
                             cell.setCellType(Cell.CELL_TYPE_STRING);  
                             cellValue = String.valueOf(cell.getRichStringCellValue().getString()) + "#" ;  
                             break;  
                         default:  
                             cellValue = "#" ;  
                     }
                     col ++;
       //              System.out.print(cellValue);  
                     rowValue += cellValue;  
                 }  
                 
                 writeSql(rowValue,map1,map2); 
         //        writeSql1(rowValue, bw, map1, map2);
       //          System.out.println(rowValue);  
                 
             }  
  //           bw.flush();  
         } catch (Exception e) {  
             e.printStackTrace();  
         } finally{  
     //        bw.close();  
         }  
    }
    
    public static void startRead1(String inpath,int beginRow,Map map1,Map map2)  throws Exception{
   	 SimpleDateFormat fmt = new SimpleDateFormat("yyyy-MM-dd");  
 //       BufferedWriter bw = new BufferedWriter(new FileWriter(new File(outpath)));  
        try {  
            // 同时支持Excel 2003、2007  
            File excelFile = new File(inpath); // 创建文件对象  
            FileInputStream is = new FileInputStream(excelFile); // 文件流  
            checkExcelVaild(excelFile);  
            Workbook workbook = getWorkbok(is,excelFile);  
            //Workbook workbook = WorkbookFactory.create(is); // 这种方式 Excel2003/2007/2010都是可以处理的  
  
            int sheetCount = workbook.getNumberOfSheets(); // Sheet的数量  
         //   System.out.println("Sheet的数量为  :" + sheetCount);
            /** 
             * 设置当前excel中sheet的下标：0开始 
             */  
            Sheet sheet = workbook.getSheetAt(0);   // 遍历第一个Sheet  
  
            // 为跳过第一行目录设置count  
            int count = beginRow;  
  
            for (Row row : sheet) {  
                // 跳过第一行的目录  
                if(count == 0){  
                    count++;  
                    continue;  
                }  
                // 如果当前行没有数据，跳出循环  
                if(row.getCell(0) == null || row.getCell(0).toString().equals("")){  
                    return ;  
                }  
                String rowValue = "";  
                for (Cell cell : row) {  
                    if(cell.toString() == null){  
                        continue;  
                    }  
                    int cellType = cell.getCellType();  
                    String cellValue = "";  
                    switch (cellType) {  
                        case Cell.CELL_TYPE_STRING:     // 文本  
                            cellValue = cell.getRichStringCellValue().getString() + "#" ;                             
                            break;  
                        case Cell.CELL_TYPE_NUMERIC:    // 数字、日期  
                            if (DateUtil.isCellDateFormatted(cell)) {  
                                cellValue = fmt.format(cell.getDateCellValue()) + "#" ;  
                            } else {  
                                cell.setCellType(Cell.CELL_TYPE_STRING);  
                                cellValue = String.valueOf(cell.getRichStringCellValue().getString()) + "#" ;  
                            }  
                            break;  
                        case Cell.CELL_TYPE_BOOLEAN:    // 布尔型  
                            cellValue = String.valueOf(cell.getBooleanCellValue()) + "#"  ;  
                            break;  
                        case Cell.CELL_TYPE_BLANK: // 空白  
                            cellValue = cell.getStringCellValue() + "#" ;  
                            break;  
                        case Cell.CELL_TYPE_ERROR: // 错误  
                            cellValue = "错误#" ;  
                            break;  
                        case Cell.CELL_TYPE_FORMULA:    // 公式  
                            // 得到对应单元格的公式  
                            //cellValue = cell.getCellFormula() + "#";  
                            // 得到对应单元格的字符串  
                            cell.setCellType(Cell.CELL_TYPE_STRING);  
                            cellValue = String.valueOf(cell.getRichStringCellValue().getString()) + "#" ;  
                            break;  
                        default:  
                            cellValue = "#" ;  
                    }  
      //              System.out.print(cellValue);  
                    rowValue += cellValue;  
                }  
          //      writeSql(rowValue,bw,map1,map2); 
                writeSql1(rowValue, map1, map2);
      //          System.out.println(rowValue);  
                
            }  
   //         bw.flush();  
        } catch (Exception e) {  
            e.printStackTrace();  
        } finally{  
 //           bw.close();  
        }  
   }
    
    
    	 public static void writeSql1(String rowValue,Map map1,Map map2) throws IOException{  
    	    	
    	       String[] sqlValue = rowValue.split("#") ;
   
    	       if(!"应结票款".equals(sqlValue[8])){
          			amount2 += Integer.parseInt(sqlValue[6]) * Integer.parseInt(sqlValue[7]);
          			
          			 int a = Integer.parseInt(sqlValue[6]);
                     int b = Integer.parseInt(sqlValue[7]);
                     int c = Integer.parseInt(sqlValue[8]);
                     
                     if(a * b != c){
                  	   System.out.println(rowValue);
                     }
                     
          		}
    	       
               map2.put(sqlValue[1], sqlValue[6]);
               
               
    	    //    sql="INSERT INTO table_name (列名1) VALUES("+ sqlValue[0].trim() + ");"+"\n";  

//    	        try {  
//    	            bw.write(sql);  
//    	            bw.newLine();  
//    	        } catch (IOException e) {  
//    	            e.printStackTrace();  
//    	        }  
    	        
    	    }  
}
