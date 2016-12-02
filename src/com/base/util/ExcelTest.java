package com.base.util;

import java.io.*;
import jxl.*;
import jxl.write.*;
import jxl.format.*;

import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.excel.Student;

import java.awt.Color;
import jxl.write.Number;
import jxl.write.Boolean;






public class ExcelTest {
	

	 public static  List list=new ArrayList();   
	
	//输出流，X轴Y轴； 自定义关联表导出
//	public static void writeExcel(OutputStream os,int x,int y,List content) throws Exception {
//		 RegUser user=new RegUser();
//		 String str="";
//		jxl.write.WritableWorkbook wwb = Workbook.createWorkbook(os);
//		jxl.write.WritableSheet ws = wwb.createSheet("Sheet", 0);
//		
////		jxl.write.Label labelC = new jxl.write.Label(x, y,content);
////		labelC = new jxl.write.Label(4, 4,content);
////		
//		if (ws != null) {
//
//			// 下面开始添加单元格
//			for (int i = 0; i < x; i++) {
//				
//				
//				for (int j = 0; j < y; j++) {
//					// 这里需要注意的是，在Excel中，第一个参数表示列，第二个表示行
//					user=(RegUser)content.get(i);
//					
//					if(j==0){
//						str=user.getUserName();
//					}else if(j==1){
//						str=user.getPassword();
//					}else if(j==2){
//						str=user.getUserContent();
//					}else if(j==3){
//						str=user.getFlag();
//					}else if(j==4){
//						str=user.getUserTime();
//					}else if(j==5){
//						str=String.valueOf(user.getId());
//					}
//					
//					
//					jxl.write.Label labelC = new Label(j, i,str);
//					try {
//						// 将生成的单元格添加到工作表中
//						ws.addCell(labelC);
//					} catch (RowsExceededException e) {
//						e.printStackTrace();
//					} catch (WriteException e) {
//						e.printStackTrace();
//					}
//
//				}
//			}
//
//		}
//
//
//
//		/*
//		 * 自定义单元格内样式；
//		 */
//		// jxl.write.WritableFont wfc = new
//		// jxl.write.WritableFont(WritableFont.ARIAL, 10, WritableFont.BOLD,
//		// false,UnderlineStyle.NO_UNDERLINE);
//		// jxl.write.WritableCellFormat wcfFC = new
//		// jxl.write.WritableCellFormat(wfc);
//		//单元格加红
//		//wcfFC.setBackground(jxl.format.Colour.RED);
//		//labelC = new jxl.write.Label(6, 0, "中国爱我 ", wcfFC);
//		
//		
//		
//		//ws.addCell(labelC);
//		// 写入Exel工作表
//		wwb.write();
//		// 关闭Excel工作薄对象
//		wwb.close();
//	
//	}
	

	
	
	
	
	
	//new file ()
	
	/**
     * 读取Excel
     *
     * @param filePath
     */
    public static void readExcel(String filePath)
    {
    	 String data2="";
    	 //DateTime tem=new DateTime("");
   		 SimpleDateFormat sdf=new SimpleDateFormat("yyyy-MM-dd");
        try
        {
            InputStream is = new FileInputStream(filePath);
            Workbook rwb = Workbook.getWorkbook(is);
            //Sheet st = rwb.getSheet("0")这里有两种方法获取sheet表,1为名字，而为下标，从0开始
            Sheet st = rwb.getSheet("Sheet1");
            int rs=st.getColumns();
            int rows=st.getRows();
            System.out.println("列数===>"+rs+"行数："+rows);
            
           
            	for(int k=0;k<rows;k++){//行
            		 for(int i=0 ;i<rs;i++){//列
            			 
                   Cell c00 = st.getCell(i,k);
            //通用的获取cell值的方式,返回字符串
            String strc00 = c00.getContents();
            //获得cell具体类型值的方式
            if(c00.getType() == CellType.LABEL)
            {
                LabelCell labelc00 = (LabelCell)c00;
                strc00 = labelc00.getString();
            }
            //excel 类型为时间类型处理;
            if(c00.getType()==CellType.DATE){
            	DateCell dc=(DateCell)c00;  
            	strc00 = sdf.format(dc.getDate());
            	
            }
            //excel 类型为数值类型处理;
            /*
            if(c00.getType()==CellType.NUMBER|| c00.getType()==CellType.NUMBER_FORMULA){
            	NumberCell nc=(NumberCell)c00; 
            	strc00=""+nc.getValue(); 
            }*/
            
            //输出
            System.out.println(">"+strc00);
            
            list.add(strc00);
            
       
   		 //列，行
//   		 data2=String.valueOf(st.getCell(1,k).getContents());
//   		 data2=data2.replace("/", "-");
//           java.util.Date dt=sdf.parse(data2);	
//           System.out.println(sdf.format(dt));   
//            	           	
            		 }
  		 System.out.println(data2+"======"+list.get(k)+"=========");	 
       }
            
            
            //关闭
            rwb.close();
        }
        catch(Exception e)
        {
            e.printStackTrace();
        }
    }

    /**
     * 输出Excel
     *
     * @param os
     */
    public static void writeExcel(OutputStream os)
    {
        try
        {
            /**
             * 只能通过API提供的工厂方法来创建Workbook，而不能使用WritableWorkbook的构造函数，
             * 因为类WritableWorkbook的构造函数为protected类型
             * method(1)直接从目标文件中读取WritableWorkbook wwb = Workbook.createWorkbook(new File(targetfile));
             * method(2)如下实例所示 将WritableWorkbook直接写入到输出流

             */
            WritableWorkbook wwb = Workbook.createWorkbook(os);
            //创建Excel工作表 指定名称和位置
            WritableSheet ws = wwb.createSheet("Test Sheet 1",0);

            //**************往工作表中添加数据*****************

            //1.添加Label对象
            Label label = new Label(0,0,"this is a label test");
            ws.addCell(label);

            //添加带有字型Formatting对象
            WritableFont wf = new WritableFont(WritableFont.TIMES,18,WritableFont.BOLD,true);
            WritableCellFormat wcf = new WritableCellFormat(wf);
            Label labelcf = new Label(1,0,"this is a label test",wcf);
            ws.addCell(labelcf);

            //添加带有字体颜色的Formatting对象
            WritableFont wfc = new WritableFont(WritableFont.ARIAL,10,WritableFont.NO_BOLD,false,
                    UnderlineStyle.NO_UNDERLINE,jxl.format.Colour.RED);
            WritableCellFormat wcfFC = new WritableCellFormat(wfc);
            Label labelCF = new Label(1,0,"This is a Label Cell",wcfFC);
            ws.addCell(labelCF);

            //2.添加Number对象
            Number labelN = new Number(0,1,3.1415926);
            ws.addCell(labelN);

            //添加带有formatting的Number对象
            NumberFormat nf = new NumberFormat("#.##");
            WritableCellFormat wcfN = new WritableCellFormat(nf);
            Number labelNF = new jxl.write.Number(1,1,3.1415926,wcfN);
            ws.addCell(labelNF);

            //3.添加Boolean对象
            Boolean labelB = new jxl.write.Boolean(0,2,false);
            ws.addCell(labelB);

            //4.添加DateTime对象
            jxl.write.DateTime labelDT = new jxl.write.DateTime(0,3,new java.util.Date());
            ws.addCell(labelDT);

            //添加带有formatting的DateFormat对象
            DateFormat df = new DateFormat("dd MM yyyy hh:mm:ss");
            WritableCellFormat wcfDF = new WritableCellFormat(df);
            DateTime labelDTF = new DateTime(1,3,new java.util.Date(),wcfDF);
            ws.addCell(labelDTF);


            //添加图片对象,jxl只支持png格式图片
            //File image = new File("d://2.png");
           // WritableImage wimage = new WritableImage(0,1,2,2,image);
           // ws.addImage(wimage);
            //写入工作表
            wwb.write();
            wwb.close();
        }
        catch(Exception e)
        {
            e.printStackTrace();
        }
    }

     
    
    /**
     * 拷贝后,进行修改,其中file1为被copy对象，file2为修改后创建的对象
     * 尽单元格原有的格式化修饰是不能去掉的，我们还是可以将新的单元格修饰加上去，
     * 以使单元格的内容以不同的形式表现
     * @param file1
     * @param file2
     */
    public static void modifyExcel(File file1,File file2)
    {
        try
        {
            Workbook rwb = Workbook.getWorkbook(file1);
            WritableWorkbook wwb = Workbook.createWorkbook(file2,rwb);//copy
            WritableSheet ws = wwb.getSheet(0);
            WritableCell wc = ws.getWritableCell(0,0);
            //判断单元格的类型,做出相应的转换
            if(wc.getType() == CellType.LABEL)
            {
                Label label = (Label)wc;
                label.setString("The value has been modified");
            }
            wwb.write();
            wwb.close();
            rwb.close();
        }
        catch(Exception e)
        {
            e.printStackTrace();
        }
    }


    //测试
    public static void main(String[] args)
    {
        try
        {
            //读Excel
        	//ExcelTest.readExcel("C:/Users/Administrator/Desktop/test23.xls");
            //输出Excel
        	/*File fileWrite = new File("C:/Users/Administrator/Desktop/test23.xls");
            fileWrite.createNewFile();
            OutputStream os = new FileOutputStream(fileWrite);
            ExcelTest.writeExcel(os);*/
            //修改Excel
          //  excelTest.modifyExcel(new File(""),new File(""));
        	
        	/*String filePath = "C:/Users/Administrator/Desktop/test23.xls";
        	InputStream is = new FileInputStream(filePath);
            Workbook rwb = Workbook.getWorkbook(is);*/
        	String filePath = "C:/Users/Administrator/Desktop/附件一.xlsx";
        	//String filePath = "C:/Users/Administrator/Desktop/test23.xls";
        	String[] columnArray = {"A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"};
        	if("xlsx".equals(filePath.split("\\.")[filePath.split("\\.").length-1])){
        		//2010
        		
        		DecimalFormat df = new DecimalFormat("0");  
                InputStream is = new FileInputStream(filePath);
                XSSFWorkbook xssfWorkbook = new XSSFWorkbook(is);
                Student student = null;
                List<Student> list = new ArrayList<Student>();
                // Read the Sheet
                for (int numSheet = 0; numSheet < xssfWorkbook.getNumberOfSheets(); numSheet++) {
                    XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(numSheet);
                    if (xssfSheet == null) {
                        continue;
                    }
                    //System.out.println(xssfSheet.getSheetName()+" 行数 "+xssfSheet.getPhysicalNumberOfRows() +"  列数"+xssfSheet.getRow(0).getPhysicalNumberOfCells());
                    //System.out.println(xssfSheet.getSheetName()+" 行数 "+xssfSheet.getPhysicalNumberOfRows() );
                    if("支付宝".equals(xssfSheet.getSheetName())){
                    	int columnNumCount = xssfSheet.getRow(0).getPhysicalNumberOfCells();
                    	int lastRowNum = xssfSheet.getLastRowNum();
                    	XSSFRow xssfRow = xssfSheet.getRow(0);
                    	for(int i=0;i<columnNumCount;i++){
                    		
                    		System.out.println("名1称---》"+xssfRow.getCell(i)+"   索引---》"+columnArray[i]);
                    	}
                    }
                    // Read the Row
                   /* for (int rowNum = 1; rowNum <= xssfSheet.getLastRowNum(); rowNum++) {
                        XSSFRow xssfRow = xssfSheet.getRow(rowNum);
                        if (xssfRow != null) {
                            student = new Student();
                            XSSFCell no = xssfRow.getCell(0);
                            XSSFCell name = xssfRow.getCell(1);
                            XSSFCell age = xssfRow.getCell(2);
                            XSSFCell score = xssfRow.getCell(3);
                            System.out.println(no+" "+name+" "+age+" "+score); 
                        }
                    }*/
                }
        	}else{
        		//2007
        		DecimalFormat df = new DecimalFormat("0");  
        		InputStream is = new FileInputStream(filePath);
        		HSSFWorkbook hssfWorkbook = new HSSFWorkbook(is);
        		Student student = null;
        		List<Student> list = new ArrayList<Student>();
        		// Read the Sheet
        		for (int numSheet = 0; numSheet < hssfWorkbook.getNumberOfSheets(); numSheet++) {
        			HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(numSheet);
        			if (hssfSheet == null) {
        				continue;
        			}else{
        				 if("支付宝".equals(hssfSheet.getSheetName())){
                         	int columnNumCount = hssfSheet.getRow(0).getPhysicalNumberOfCells();
                         	int lastRowNum = hssfSheet.getLastRowNum();
                         	HSSFRow hssfRow = hssfSheet.getRow(0);
                         	for(int i=0;i<columnNumCount;i++){
                         		
                         		System.out.println("名称---》"+hssfRow.getCell(i)+"   索引---》"+columnArray[i]);
                         	}
                         }
        				/*for (int rowNum = 1; rowNum <= hssfSheet.getLastRowNum(); rowNum++) {
        					HSSFRow hssfRow = hssfSheet.getRow(rowNum);
        					if (hssfRow != null) {
        						student = new Student();
        						HSSFCell no = hssfRow.getCell(0);
        						HSSFCell name = hssfRow.getCell(1);
        						HSSFCell age = hssfRow.getCell(2);
        						HSSFCell score = hssfRow.getCell(3);
        						System.out.println(no+" "+name+" "+age+" "+score); 
        					}
        				}*/
        			}
        			
        		}
        	}
        	}
        catch(Exception e)
        {
           e.printStackTrace();
        }
    }

	

}
