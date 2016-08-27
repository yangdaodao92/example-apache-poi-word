package word.template;

import org.apache.poi.xwpf.usermodel.*;

import java.io.*;
import java.net.URISyntaxException;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class WordTemplate {

    public static void main(String[] args) throws URISyntaxException, IOException {
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
		XWPFDocument document = w.extractTemplate(new WordTemplate().load("两周计划.docx"), properties);
		FileOutputStream os = new FileOutputStream("两周计划2.docx");
		document.write(os);

    }

    private InputStream load(String path){
		InputStream stream;
		try {
			//寻找包外配置
			stream = new FileInputStream(System.getProperty("user.dir") + System.getProperty("file.separator") + path);
		} catch (FileNotFoundException e) {
			//寻找包内配置
			stream = getClass().getClassLoader().getResourceAsStream(path);
		}
		return stream;
	}

    /**
     * Extrai o template do word
     * @param stream inputstream do arquivo word
     * @param properties chave/valor a ser usado.
     * @return XWPFDocument
     * @throws IOException
     */
    public XWPFDocument extractTemplate(InputStream stream, Map<String, String> properties) throws IOException {
        XWPFDocument document = new XWPFDocument(stream);
        replaceParagraphs(document.getParagraphs(), properties);
        replaceTables(document.getTablesIterator(), properties);
        return document;
    }

    /**
     * Altera os valores dos paragrafos.
     * @param paragraphs paragrafos
     * @param properties propriedades
     */
    private void replaceParagraphs(List<XWPFParagraph> paragraphs, Map<String, String> properties) {
        for (XWPFParagraph paragraph : paragraphs) {
        	replaceInPara(paragraph, properties);
        }
    }

	public void replaceInPara(XWPFParagraph paragraph, Map<String, String> properties) {
		int start = -1;
		String str = "";
		if (matcher(paragraph.getParagraphText()).find()) {
			List<XWPFRun> runs = paragraph.getRuns();
			for (int i = 0; i < runs.size(); i++) {
				String runText = runs.get(i).toString();
				if (runText.startsWith("#")) start = i;
				if (start != -1) str += runText;
				if (runText.endsWith("}") && start != -1) {
					for (int j = 0; j < i - start; j++) {
						paragraph.removeRun(start + 1);
					}
					for (Entry<String, String> entry : properties.entrySet()) {
						if (str.equals(entry.getKey())) {
							runs.get(start).setText(properties.get(entry.getKey()), 0);
							break;
						}
					}
					break;
				}
			}
			replaceInPara(paragraph, properties);
		}
	}

    /**
     * Altera os valores da Table
     * @param itTable table
     * @param properties proprierties
     */
    private void replaceTables(Iterator<XWPFTable> itTable, Map<String, String> properties) {
        while (itTable.hasNext()) {
            XWPFTable table = itTable.next();
            extractLines(properties, table);
        }
    }

    /**
     * Altera os valores das linhas da tabela.
     * @param properties propriedades
     * @param table tabela
     */
    private void extractLines(Map<String, String> properties, XWPFTable table) {
		List<XWPFTableCell> cells;
		List<XWPFTableRow> rows = table.getRows();
		for (XWPFTableRow row : rows) {
			cells = row.getTableCells();
			for (XWPFTableCell cell : cells) {
				replaceParagraphs(cell.getParagraphs(), properties);
			}
		}
    }

	private Matcher matcher(String str) {
		Pattern pattern = Pattern.compile("#\\{.+?\\}", Pattern.CASE_INSENSITIVE);
		return pattern.matcher(str);
	}
}
