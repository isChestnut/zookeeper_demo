import java.io.File;
import java.io.IOException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;
import java.util.UUID;

import com.google.zxing.WriterException;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

public class ExcelUtils {
	/**
     * 获取属性值
     * @param fieldName 字段名称
     * @param o 对象
     * @return  Object
     */
   private static Object getFieldValueByName(String fieldName, Object o) {
       try {
           String firstLetter = fieldName.substring(0, 1).toUpperCase();    
           String getter = "get" + firstLetter + fieldName.substring(1);    //获取方法名
           Method method = o.getClass().getMethod(getter, new Class[] {});  //获取方法对象
           Object value = method.invoke(o, new Object[] {});    //用invoke调用此对象的get字段方法
               return value;  //返回值
       } catch (Exception e) {
           e.printStackTrace();
           return null;    
       }    
   }

   
   /**
    * 将list集合转成Excel文件
    * @param list  对象集合
    * @param path  输出路径
    * @return   返回文件路径
    */
   public static String createExcel(ArrayList<List>  list,String path) throws WriterException{
       String result = "";
       if(list.size()==0||list==null){
           result = "没有对象信息";
       }else{
    /*       Object o = list.get(0);
           Class<? extends Object> clazz = o.getClass();
           String className = clazz.getSimpleName();
           Field[] fields=clazz.getDeclaredFields();  */  //这里通过反射获取字段数组
           File folder = new File(path);
           if(!folder.exists()){
               folder.mkdirs();
           }
           String fileName = "完整数据";
           String name = fileName.concat(".xls");
           WritableWorkbook book = null;
           File file = null;
           try {
               file = new File(path.concat(File.separator).concat(name));
               book = Workbook.createWorkbook(file);  //创建xls文件
               WritableSheet sheet  =  book.createSheet("sheet1",0);
             for (int j = 0; j < list.size(); j++) {
                   for (int k = 0; k < list.get(j).size(); k++) {
                	   String value="";
                	   if("null".equalsIgnoreCase(String.valueOf(list.get(j).get(k)))){
                		   value="";
                	   }else{
                		   value = String.valueOf(list.get(j).get(k));
                	   }
                	   sheet.addCell(new Label(k,j,value));
				}
                  
             }
               book.write();
               result = file.getPath();
           } catch (Exception e) {
               // TODO Auto-generated catch block
               result = "SystemException";
               e.printStackTrace();
           }finally{
               fileName = null;
               name = null;
               folder = null;
               file = null;
               if(book!=null){
                   try {
                       try {
						book.close();
					} catch (WriteException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
                   } catch (IOException e) {
                       // TODO Auto-generated catch block
                       result = "IOException";
                       e.printStackTrace();
                   }
               }
           }

       }

       return result;   //最后输出文件路径
   }

}
