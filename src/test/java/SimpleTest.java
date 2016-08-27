import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.Test;
import word.template.WordTemplate;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.util.HashMap;
import java.util.Map;

public class SimpleTest {
	@Test
	public void testTemplate() throws IOException {
		Map<DayOfWeek, String> map = new HashMap<>();
		map.put(DayOfWeek.MONDAY, "星期一");
		map.put(DayOfWeek.TUESDAY, "星期二");
		map.put(DayOfWeek.WEDNESDAY, "星期三");
		map.put(DayOfWeek.THURSDAY, "星期四");
		map.put(DayOfWeek.FRIDAY, "星期五");
		map.put(DayOfWeek.SATURDAY, "星期六");
		map.put(DayOfWeek.SUNDAY, "星期日");

		LocalDate local1 = LocalDate.now();
		LocalDate local2 = local1.plusDays(10);

		String file = "src/main/resources/两周计划.docx";
		FileInputStream in = new FileInputStream(file);
		WordTemplate w = new WordTemplate();
		Map<String, String> properties = new HashMap<>();
		properties.put("#{year1}", local1.getYear() + "");
		properties.put("#{month1}", local1.getMonthValue() + "");
		properties.put("#{day1}", local1.getDayOfMonth() + "");
		properties.put("#{year2}", local2.getYear() + "");
		properties.put("#{month2}", local2.getMonthValue() + "");
		properties.put("#{day2}", local2.getDayOfMonth() + "");
		for (int i=1; i<=20; i++){
			LocalDate localNow = local1.plusDays(i-1);
			properties.put("#{week"+i+"}", localNow.getDayOfMonth() + map.get(localNow.getDayOfWeek()));
		}
		XWPFDocument document = w.extractTemplate(in, properties);
		FileOutputStream os = new FileOutputStream("src/main/resources/两周计划2.docx");
		document.write(os);
	}
}
