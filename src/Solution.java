import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Solution {
	// ʹ��POI����excel������
	public static void createWorkBook() throws IOException {
		// ����excel������
		Workbook wb = new HSSFWorkbook();
		// ������һ��sheet��ҳ��������Ϊ new sheet
		Sheet sheet = wb.createSheet("new sheet");
		// Row ��
		// Cell ����
		// Row �� Cell ���Ǵ�0��ʼ������

		// ����һ�У���ҳsheet��
		Row row = sheet.createRow((short) 0);
		// ��row���ϴ���һ������
		Cell cell = row.createCell(0);
		// ���÷������ʾ
		cell.setCellValue(1);

		// Or do it on one line.
		row.createCell(1).setCellValue(1.2);
		row.createCell(2).setCellValue("This is a string �ٶȷ�������");
		row.createCell(3).setCellValue(true);

		// ����һ���ļ� ����Ϊworkbook.xls
		FileOutputStream fileOut = new FileOutputStream("workbook.xls");
		// �����洴���Ĺ�����������ļ���
		wb.write(fileOut);
		// �ر������
		fileOut.close();
	}

	// ʹ��POI����excel�������ļ�
	public static void readWorkBook() throws Exception {
		// poi��ȡexcel
		// ����Ҫ������ļ���������
		InputStream inp = new FileInputStream("D:\\tmp\\8�·ݿ�����ϸ��2.xls");

		// �������������������� ��������������
		Workbook wb = WorkbookFactory.create(inp);
		// �õ���һҳ sheet
		// ҳSheet�Ǵ�0��ʼ������
		Sheet sheet = wb.getSheetAt(0);
		// �ڵ�һ�����ҵ��ϰ�ʱ����°�ʱ�������
		int startWorkTimeIndex = -1;
		int stopWorkTimeIndex = -1;
		int dateIndex = -1;
		for (Cell cell : sheet.getRow(0)) {
			if (cell.getStringCellValue().equals("�ϰ�ʱ��")) {
				startWorkTimeIndex = cell.getColumnIndex();
			} else if (cell.getStringCellValue().equals("�°�ʱ��")) {
				stopWorkTimeIndex = cell.getColumnIndex();
			} else if (cell.getStringCellValue().equals("����")) {
				dateIndex = cell.getColumnIndex();
			}
		}
		if (startWorkTimeIndex == -1 || stopWorkTimeIndex == -1) {
			System.out.println("û���ҵ��ϰ�ʱ������°�ʱ��");
			return;
		}
		SimpleDateFormat sdf = new SimpleDateFormat("HH:mm:ss");
		SimpleDateFormat sdf2 = new SimpleDateFormat("yyyy-MM-dd");
		// ����foreachѭ�� ����sheet�е�������
		Calendar c = Calendar.getInstance();
		for (int i = 1; i < sheet.getLastRowNum(); ++i) {
			// ����row�е����з���
			long startTime = 0;
			long endTime = 0;
			long date = 0;
			String name = null;
			boolean isSundayOrSaturday = false;
			for (Cell cell : sheet.getRow(i)) {
				if (cell.getColumnIndex() == dateIndex) {
					date = sdf2.parse(cell.getStringCellValue()).getTime();
					c.setTimeInMillis(date);
					int week = c.get(Calendar.DAY_OF_WEEK);
					if (week == Calendar.SUNDAY || week == Calendar.SATURDAY) {
						isSundayOrSaturday = true;
						break;
					}
				}
				if (cell.getColumnIndex() == startWorkTimeIndex) {
					if (!"".equals(cell.getStringCellValue().trim())) {
						startTime = sdf.parse(cell.getStringCellValue()).getTime();
					}
				} else if (cell.getColumnIndex() == stopWorkTimeIndex) {
					if (!"".equals(cell.getStringCellValue().trim())) {
						endTime = sdf.parse(cell.getStringCellValue()).getTime();
					}
				} else if (cell.getColumnIndex() == 1) {
					name = cell.getStringCellValue();
				}
			}
			if (isSundayOrSaturday)
				continue;
			System.out.print(name + " " + sdf2.format(new Date(date)) + " " + getDayOfWeek(c));
			if (startTime == 0) {
				System.out.print("->�ϰ�δ��");
			} else {
				System.out.print("->�ϰ�����");
			}
			if (endTime == 0) {
				System.out.print("->�°�δ��");
			} else {
				System.out.print("->�°�����");
			}
//			else {
				long duration = endTime - startTime;
				// ÿһ�������֮����
				if (startTime != 0 && endTime != 0) {
					if ( duration >= (8 * 3600 * 1000)) {
						//System.out.print("<����>");
					} else {
						System.out.print("<�ٵ�>");
					}
				}
//			}
			System.out.println();
		}
		// �ر�������
		inp.close();
	}

	public static String getDayOfWeek(Calendar c) {
		int dayOfWeek = c.get(Calendar.DAY_OF_WEEK);
		if (dayOfWeek == Calendar.MONDAY) {
			return "����һ";
		}
		if (dayOfWeek == Calendar.TUESDAY) {
			return "���ڶ�";
		}
		if (dayOfWeek == Calendar.WEDNESDAY) {
			return "������";
		}
		if (dayOfWeek == Calendar.THURSDAY) {
			return "������";
		}
		if (dayOfWeek == Calendar.FRIDAY) {
			return "������";
		}
		if (dayOfWeek == Calendar.SATURDAY) {
			return "������";
		}
		return "������";
	}

	public static void main(String[] args) {
		try {
			Solution.readWorkBook();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
