import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.lang.StringUtils;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.zxing.WriterException;

public class RWExcel {
	private String filePath;
	private String anotherfilePath;

	/**
	 * 构造方法
	 */

	public RWExcel(String filePath, String anotherfilePath) {

		this.filePath = filePath;
		this.anotherfilePath = anotherfilePath;
	}

	/**
	 * 
	 * 读取excel 封装成集合 该程序需要传入一份excel 和excel的列数 行数由程序自动检测 注意：该方法统计的行数是默认第一行为title
	 * 不纳入统计的
	 * 
	 * @return
	 * 
	 */
	// @Test
	public ArrayList<List> ReadExcel(int sheetNum) {

		// int column = 5;//column表示excel的列数

		ArrayList<List> list = new ArrayList<List>();

		try {
			// 建需要读取的excel文件写入stream
			XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(filePath));
			// 指向sheet下标为0的sheet 即第一个sheet 也可以按在sheet的名称来寻找
			XSSFSheet sheet = workbook.getSheetAt(sheetNum);
			// 获取sheet1中的总行数
			int rowTotalCount = sheet.getLastRowNum();

			// 获取总列数
			int columnCount = sheet.getRow(0).getPhysicalNumberOfCells();

			// System.out.println("行数为："+rowTotalCount+"列数为："+columnCount);

			for (int i = 0; i <= rowTotalCount; i++) {
				// 获取第i列的row对象
				XSSFRow row = sheet.getRow(i);

				ArrayList<String> listRow = new ArrayList<String>();

				for (int j = 0; j < columnCount; j++) {
					// 下列步骤为判断cell读取到的数据是否为null 如果不做处理 程序会报错
					String cell = null;
					// 如果未null则加上""组装成非null的字符串
					System.out.println("行"+ i+"列"+j);
					if (row.getCell(j) == null) {

						cell = row.getCell(j) + "";

						listRow.add(cell);
						// 如果读取到额cell不为null 则直接加入 listRow集合
					} else {
						cell = row.getCell(j).toString();
						listRow.add(cell);
					}
					// 在第i列 依次获取第i列的第j个位置上的值 %15s表示前后间隔15个字节输出
					// System.out.printf("%15s", cell);

				}

				list.add(listRow);

				// System.out.println();
			}

		} catch (FileNotFoundException e) {

			e.printStackTrace();
		} catch (IOException e) {

			e.printStackTrace();
		}

		return list;
	}

	/**
	 * 读取另外一份Excel 保存成list集合
	 */

	public ArrayList<List> ReadAnotherExcel(int anotherSheetNum) {

		ArrayList<List> list = new ArrayList<List>();

		try {
			// 建需要读取的excel文件写入stream
			XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(anotherfilePath));
			// 指向sheet下标为0的sheet 即第一个sheet 也可以按在sheet的名称来寻找
			XSSFSheet sheet = workbook.getSheetAt(anotherSheetNum);
			// 获取sheet1中的总行数
			int rowTotalCount = sheet.getLastRowNum();
			// 获取总列数
			int columnCount = sheet.getRow(0).getPhysicalNumberOfCells();

			// System.out.println("行数为："+rowTotalCount+"列数为："+columnCount);

			for (int i = 0; i <= rowTotalCount; i++) {
				// 获取第i列的row对象
				XSSFRow row = sheet.getRow(i);

				ArrayList<String> listRow = new ArrayList<String>();

				for (int j = 0; j < columnCount; j++) {
					// 下列步骤为判断cell读取到的数据是否为null 如果不做处理 程序会报错
					String cell = null;
					// 如果未null则加上""组装成非null的字符串
                    System.out.println( "^o^ ^o^ ^o^ ^o^" +i+ "^o^ ^o^ ^o^ ^o^"+j);
					if (row.getCell(j) == null) {
						cell = row.getCell(j) + "";

						listRow.add(cell);
						// 如果读取到额cell不为null 则直接加入 listRow集合
					} else {
						cell = row.getCell(j).toString();
						listRow.add(cell);
					}
					// 在第i列 依次获取第i列的第j个位置上的值 %15s表示前后间隔15个字节输出
					// System.out.printf("%15s", cell);

				}

				list.add(listRow);

				// System.out.println();
			}

		} catch (FileNotFoundException e) {

			e.printStackTrace();
		} catch (IOException e) {

			e.printStackTrace();
		}

		return list;
	}

	/**
	 * 调试方法
	 */

	public static void main(String[] args) {

		RWExcel excel = new RWExcel("C:\\Users\\Administrator.A6EB3XSRUSVBQ23\\Desktop\\temp20191202\\11.30个贷台账.xlsx","C:\\Users\\Administrator.A6EB3XSRUSVBQ23\\Desktop\\temp20191202\\1202放款业务统计.xlsx");

		ArrayList<List> list1 = excel.ReadExcel(0);

		ArrayList<List> list2 = excel.ReadAnotherExcel(0);

		System.out.println("==========================");

		
		ArrayList<List> list = excel.getFinally(list1,list2,19,1);
		try {
			ExcelUtils.createExcel(list, "C:\\Users\\Administrator.A6EB3XSRUSVBQ23\\Desktop\\temp20191202");
			System.out.println(" ^o^ ^o^ ^o^ ^o^ ^o^ ^o^ ^o^ ^o^ ^o^ ^o^ ^o^ ^o^ ^o^ ^o^ ");
		} catch (WriterException e) {
			e.printStackTrace();
		}
	}

	//合并一样的数据
	public ArrayList<List> getFinally(ArrayList<List> list1, ArrayList<List> list2, int one, int two) {
		//匹配完成数据
		ArrayList<List> list4 = list1;
		
		// 取出表格二中待匹配的元素
		List<Object> list3 = new ArrayList<>();
		for (int i = 0; i < list2.size(); i++) {
			list3.add(list2.get(i).get(two));
		}
		// 循环表一中的数据到表二中匹配
		for (int i = 0; i < list1.size(); i++) {
			if(i==0){
				list4.get(i).addAll(list2.get(i));
				continue;
			}
			// 表格一中与表格二需匹配的元素 a
			Object a = list1.get(i).get(one);
			//表一缺失不匹配
			if (StringUtils.isBlank(String.valueOf(a))) {
				continue;
			}
			if (list3.contains(a)) {
				// 获取表格一中数据在表格二中的位置 b
				int b = list3.indexOf(a);
				// 将匹配到的数据添加至表格一
					list4.get(i).addAll(list2.get(b));
			}
		}
		return list4;
	}
	
	//检验不一样的数据
		public ArrayList<List> findNotSame(ArrayList<List> list1, ArrayList<List> list2, int one, int two) {
			//匹配完成数据
			ArrayList<List> list4 = new ArrayList<>();
			
			// 取出表格二中待匹配的元素
			List<Object> list3 = new ArrayList<>();
			for (int i = 0; i < list2.size(); i++) {
				list3.add(list2.get(i).get(two));
			}
			// 循环表一中的数据到表二中匹配
			for (int i = 0; i < list1.size(); i++) {
			/*	if(i==0){
					list4.add((List)list1.get(i));
					continue;
				}*/
				// 表格一中与表格二需匹配的元素 a
				Object a = list1.get(i).get(one);
				//表一缺失不匹配
				if (StringUtils.isBlank(String.valueOf(a))) {
					continue;
				}
				if (!list3.contains(a)) {
			/*		// 获取表格一中数据在表格二中的位置 b
					int b = list3.indexOf(a);
					// 将匹配到的数据添加至表格一
						//list4.get(i).addAll(list2.get(b));
					list4.add((List)list1.get(i));
					list4.get(list4.size()-1).add(list2.get(b).get(2));*/
					list4.add((List)list1.get(i));
					
				}
			}
			return list4;
		}
		public ArrayList<Object> findSame(ArrayList<List> list){
			ArrayList<Object> list3 = new ArrayList<>();
			ArrayList<Object> list4 = new ArrayList<>();
			for (int i = 0; i < list.size(); i++) {
				list.get(i).get(0);
				if(!list4.contains(list.get(i).get(0))){
					list4.add(list.get(i).get(0));
				}else{
					list3.add(list.get(i).get(0));
				}
			}
			
			for (int i = 0; i < list3.size(); i++) {
				System.out.println(list3.get(i));
			}
			return list3;
		}
}
