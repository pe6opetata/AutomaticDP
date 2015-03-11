package classes;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;

@SuppressWarnings("unused")
public class ProcessExcelFile {

	private static final String OSI = "osi.axs";
	private static final String LABELS = "label.lab";
	private static final int SIZE_OFF = 4;

	private File file;
	private FileInputStream inputFile;
	private StringBuilder logBuffer;

	public ProcessExcelFile() {
		file = null;
		inputFile = null;
		logBuffer = new StringBuilder();
	}

	public ProcessExcelFile(String fileName) throws IOException,
			FileNotFoundException {
		file = new File(fileName);
		inputFile = new FileInputStream(file);
		logBuffer = new StringBuilder();
	}

	public void process() throws IOException {

		XSSFWorkbook workbook = new XSSFWorkbook(inputFile);

		XSSFSheet questions = workbook.getSheet("Questions");
		XSSFSheet answers = workbook.getSheet("Answers 1");

		Iterator<Row> questionRowIterator = questions.iterator();

		HashMap<String, Question> allQuestions = new LinkedHashMap<String, Question>();

		logInfo("Започвам обработката на файл " + file.getAbsolutePath()
				+ "...");

		// Adding all questions in the hashmap

		while (questionRowIterator.hasNext()) {
			Row row = questionRowIterator.next();
			Cell qstID = row.getCell(1);
			Cell label = row.getCell(2);
			Cell code = row.getCell(4);
			String questionID = getPrefix(qstID);

			Question question = allQuestions.get(questionID);
			if (question == null) {
				question = new Question();
				String clearedLabel = Jsoup.parse(label.toString()).text();
				question.setLabel(clearedLabel);
				question.setID(questionID);
				logInfo("I'm adding question " + questionID + " : "
						+ clearedLabel);
			} else {
				Answer answer = new Answer();
				answer.setCode((int) code.getNumericCellValue());
				if (label.getCellType() == Cell.CELL_TYPE_NUMERIC)
					answer.setLabel(Integer.toString((int) label
							.getNumericCellValue()));
				else {
					String clearedLabel = Jsoup.parse(label.toString()).text();
					answer.setLabel(clearedLabel);
				}

				question.addAnswer(answer);
				logInfo("I'm adding answer to " + questionID + " : "
						+ answer.getLabel());
			}
			allQuestions.put(questionID, question);
		}

		Iterator<Cell> qstIDIterator = answers.getRow(0).cellIterator();
		Iterator<Cell> typeIterator = answers.getRow(1).cellIterator();
		Iterator<Cell> coloneIterator = answers.getRow(2).cellIterator();
		Iterator<Cell> qstIterator = answers.getRow(3).cellIterator();

		while (typeIterator.hasNext()) {
			Cell questionIDCell = qstIDIterator.next();
			Cell typeCell = typeIterator.next();
			Cell columnCell = coloneIterator.next();
			Cell questionCell = qstIterator.next();
			String questionID = getPrefix(questionIDCell);
			String questionTitle = getPrefix(questionCell);

			Question question = allQuestions.get(questionID);
			if (question != null) {
				if (question.getStartCol() == 0) {
					question.setStartCol((int) columnCell.getNumericCellValue());
					question.setEndCol((int) columnCell.getNumericCellValue()
							+ SIZE_OFF - 1);
					question.setType(typeCell.toString());
					question.setTitle(questionTitle);
					logInfo("I'm configuring question " + questionID
							+ " right now!");
				}
			} else {
				System.err.println("Unable to find the question with ID: "
						+ questionID + "!");
			}
		}

		FileOutputStream osi = new FileOutputStream(new File(file.getParent()
				+ "//" + OSI));
		// FileOutputStream tables = new FileOutputStream(new File(TABLES));
		FileOutputStream label = new FileOutputStream(new File(file.getParent()
				+ "//" + LABELS));

		BufferedWriter writeToOsi = new BufferedWriter(new OutputStreamWriter(
				osi, "UTF8"));
		// BufferedWriter writeToTables = new BufferedWriter(new
		// OutputStreamWriter(tables));
		BufferedWriter writeToLabel = new BufferedWriter(
				new OutputStreamWriter(label, "UTF8"));

		startLabel(writeToLabel);

		for (Question question : allQuestions.values()) {
			if (question.getAnswers() != null
					&& !question.getAnswers().isEmpty()) {
				makeOsi(writeToOsi, question);
				makeLabels(writeToLabel, question);
			}
		}

		endLabel(writeToLabel);
	}

	private void startLabel(BufferedWriter output) throws IOException {
		output.write("CROSS ANALYSIS OF BREAKS");
		output.newLine();
		output.flush();
		output.newLine();
		output.flush();
		output.write("TWOSHAPES &#0025; THOSE RATING THE IDEA AS ONE TO SELL OR DOUBLE SHARES IN ");
		output.newLine();
		output.flush();
		output.write("IDEAS ALL SEEING EACH IDEA");
		output.newLine();
		output.flush();
		output.write("SELLSHARE THOSE RATING IDEA AS ONE TO SELL SHARE IN");
		output.newLine();
		output.flush();
		output.write("DOUBLESHARE THOSE RATING IDEA AS ONE TO DOUBLE SHARE IN");
		output.newLine();
		output.flush();
		output.write("CONCEPT CONCEPT MOST PREFERED");
		output.newLine();
		output.flush();
		output.write("FACE_TT FACE TRACE SUMMARY");
		output.newLine();
		output.flush();
		output.write("SUMMARY_PRODUCT1 PM Autocharter");
		output.newLine();
		output.flush();
		output.write("SUMMARY_PRODUCT2 PM StarCharter");
		output.newLine();
		output.flush();
		output.write("IDEASEEN ALL SEEING EACH IDEA");
		output.newLine();
		output.flush();
		output.newLine();
		output.flush();
		output.newLine();
		output.flush();
	}

	private void endLabel(BufferedWriter output) throws IOException {
		output.write("TOT TOTAL");
		output.newLine();
		output.flush();
		output.write("T_W TOTAL");
		output.newLine();
		output.flush();
		output.write("TOTAL Base: All respondents");
		output.newLine();
		output.flush();
		output.write("T_NW2 Base: All respondents");
		output.newLine();
		output.flush();
		output.write("T_FACE Base: All ratings");
		output.newLine();
		output.flush();
		output.write("T_TWO Base: Two ideas rated");
		output.newLine();
		output.flush();
		output.write("T_SEEN Base: All seeing each idea");
		output.newLine();
		output.flush();
		output.write("T_FT2_1 Base: All who feel emotion: Surprise");
		output.newLine();
		output.flush();
		output.write("T_FT2_2 Base: All who feel emotion: Happiness");
		output.newLine();
		output.flush();
		output.write("T_FT2_4 Base: All who feel emotion: Sadness");
		output.newLine();
		output.flush();
		output.write("T_FT2_5 Base: All who feel emotion: Fear");
		output.newLine();
		output.flush();
		output.write("T_FT2_6 Base: All who feel emotion: Anger");
		output.newLine();
		output.flush();
		output.write("T_FT2_7 Base: All who feel emotion: Disgust");
		output.newLine();
		output.flush();
		output.write("T_FT2_8 Base: All who feel emotion: Contempt");
		output.newLine();
		output.flush();
	}

	private void makeOsi(BufferedWriter output, Question question)
			throws IOException {
		if (isQuestionOfType("single", question)) {
			output.write("l " + question.getTitle());
			output.newLine();
			output.flush();
			output.write("ttl " + question.getTitle());
			output.newLine();
			output.flush();
			output.write("ttl TOTAL");
			output.newLine();
			output.flush();
			output.write("C n10  T_NW   ;wm=0");
			output.newLine();
			output.flush();
			output.write("n10 TOT");
			output.newLine();
			output.flush();
			int i = 1;
			for (Answer answer : question.getAnswers()) {
				output.write("n01   " + i + "   ;c=c(" + question.getStartCol()
						+ "," + question.getEndCol() + ") .in. ("
						+ answer.getCode() + ")");
				output.newLine();
				output.flush();
				i++;
			}
			output.newLine();
			output.flush();
			output.newLine();
			output.flush();

		} else if (isQuestionOfType("multi", question)) {
			output.write("l " + question.getTitle());
			output.newLine();
			output.flush();
			output.write("ttl " + question.getTitle());
			output.newLine();
			output.flush();
			output.write("ttl TOTAL");
			output.newLine();
			output.flush();
			output.write("C n10  T_NW   ;wm=0");
			output.newLine();
			output.flush();
			output.write("n10 TOT");
			output.newLine();
			output.flush();
			int i = 1;
			for (Answer answer : question.getAnswers()) {
				// int start = question.getStartCol() +
				// ProcessExcelFile.SIZE_OFF * (i - 1);
				int end = question.getEndCol() + ProcessExcelFile.SIZE_OFF
						* (i - 1);
				output.write("n01   " + i + "   ;c=c" + end + "'1'");
				output.newLine();
				output.flush();
				i++;
			}
			if (question.getTitle().substring(1).length() == 1) {
				output.write("n25   ;inc=t10"
						+ question.getTitle().substring(1) + ";c=t10"
						+ question.getTitle().substring(1) + " .gt. 0");
				output.newLine();
				output.flush();
			} else {
				output.write("n25   ;inc=t1" + question.getTitle().substring(1)
						+ ";c=t1" + question.getTitle().substring(1)
						+ " .gt. 0");
				output.newLine();
				output.flush();
			}
			output.write("n12   _99   ;dec=6");
			output.newLine();
			output.flush();
			output.write("n20   _98   ;dec=6");
			output.newLine();
			output.flush();
			output.write("n19   _97   ;dec=6");
			output.newLine();
			output.flush();
			output.newLine();
			output.flush();
			output.newLine();
			output.flush();
		} else if (isQuestionOfType("facetrace", question)) {
			output.write("#include ftr.qin;col(a)="
					+ (question.getEndCol() + 32) + ";col(b)="
					+ question.getEndCol() + ";lvl=prod");
			output.newLine();
			output.flush();
			output.newLine();
			output.flush();
			output.newLine();
			output.flush();
			// Display proper warning
			// to include ftr.qin
		} else if (isQuestionOfType("srat", question)) {
			output.write("#include s_rat" + question.getAnswers().size()
					+ ".qin;col(b)=" + (question.getEndCol() - 1) + ";qst="
					+ question.getTitle() + ";lvl=prod");
			output.newLine();
			output.flush();
			output.newLine();
			output.flush();
			output.newLine();
			output.flush();
			// Display proper warning
			// to include s_rat(n).qin
		} else if (isQuestionOfType("+2-2", question)
				|| isQuestionOfType("-2+2", question)) {
			output.write("#include +2-2_" + question.getAnswers().size()
					+ ".qin;col(b)=" + (question.getEndCol() - 1) + ";qst="
					+ question.getTitle() + ";lvl=prod");
			output.newLine();
			output.flush();
			output.newLine();
			output.flush();
			output.newLine();
			output.flush();
			// Display proper warning
			// to include +2-2_(n).qin
		}

	}

	private void makeLabels(BufferedWriter output, Question question)
			throws IOException {
		if (isQuestionOfType("srat", question)) {
			int i = 1;
			output.write(question.getTitle().toUpperCase() + " "
					+ question.getTitle().toUpperCase()
					+ ". STANDARD RATINGS: MEANS SUMMARY - "
					+ question.getLabel());
			output.newLine();
			output.flush();
			for (Answer answer : question.getAnswers()) {
				output.write("   _" + i + " " + answer.getLabel());
				output.newLine();
				output.flush();
				i++;
			}
			output.newLine();
			output.flush();

			i = 1;

			output.write(question.getTitle().toUpperCase() + "I &#0025; "
					+ question.getTitle().toUpperCase()
					+ "(i). STANDARD RATINGS: TOP BOX - " + question.getLabel());
			output.newLine();
			output.flush();
			output.write(question.getTitle().toUpperCase() + "II "
					+ question.getTitle().toUpperCase()
					+ "(ii). STANDARD RATINGS: SECOND BOX - "
					+ question.getLabel());
			output.newLine();
			output.flush();
			output.write(question.getTitle().toUpperCase() + "III "
					+ question.getTitle().toUpperCase()
					+ "(iii). STANDARD RATINGS: TOP 2 BOX - "
					+ question.getLabel());
			output.newLine();
			output.flush();
			for (Answer answer : question.getAnswers()) {
				output.write("   " + i + " " + answer.getLabel());
				output.newLine();
				output.flush();
				i++;
			}
			output.write("   99 None of these");
			output.newLine();
			output.flush();
			output.newLine();
			output.flush();

			i = 1;

			output.write("S_" + question.getTitle().toUpperCase() + " &#0025; "
					+ "SUMMARY:");
			output.newLine();
			output.flush();
			for (Answer answer : question.getAnswers()) {
				output.write(question.getTitle().toUpperCase() + "_" + i + " "
						+ question.getTitle().toUpperCase() + "(" + i
						+ "). STANDARD RATINGS: " + answer.getLabel());
				output.newLine();
				output.flush();
				i++;
			}
			output.write("  1 -3 Very negative (-3.)");
			output.newLine();
			output.flush();
			output.write("  2 -2 .. (-2.)");
			output.newLine();
			output.flush();
			output.write("  3 -1 .. (-1.)");
			output.newLine();
			output.flush();
			output.write("  4 0 Neutral (0.)");
			output.newLine();
			output.flush();
			output.write("  5 +1 .. (1.)");
			output.newLine();
			output.flush();
			output.write("  6 +2 .. (2.)");
			output.newLine();
			output.flush();
			output.write("  7 +3 Very positive (3.)");
			output.newLine();
			output.flush();
			output.write("  77 Top box");
			output.newLine();
			output.flush();
			output.write("  67 Top 2 box");
			output.newLine();
			output.flush();
			output.write("  345 Middle 3 box");
			output.newLine();
			output.flush();
			output.write("  12 Bottom 2 box");
			output.newLine();
			output.flush();
			output.write("  15 Other");
			output.newLine();
			output.flush();
			output.write("  _99 Mean");
			output.newLine();
			output.flush();
			output.write("  _98 Error Variance");
			output.newLine();
			output.flush();
			output.write("  _97 Standard Error");
			output.newLine();
			output.flush();
			output.newLine();
			output.flush();
			output.newLine();
			output.flush();
		} else if (isQuestionOfType("facetrace", question)) {
			output.write("C Things to be edited: FT1, FT2, <specify>, <question_name>, <question_name2>");
			output.newLine();
			output.flush();
			output.newLine();
			output.flush();
			output.write("FT1S <question_name2>. SUMMARY MEAN: DEGREE TO WHICH PEOPLE WOULD FEEL EMOTIONS (3=Strongly, 1=Not very strongly)");
			output.newLine();
			output.flush();
			output.write("  _1 Surprise");
			output.newLine();
			output.flush();
			output.write("  _2 Happy");
			output.newLine();
			output.flush();
			output.write("  _3 Neutral");
			output.newLine();
			output.flush();
			output.write("  _4 Sadness");
			output.newLine();
			output.flush();
			output.write("  _5 Fear");
			output.newLine();
			output.flush();
			output.write("  _6 Anger");
			output.newLine();
			output.flush();
			output.write("  _7 Disgust");
			output.newLine();
			output.flush();
			output.write("  _8 Contempt");
			output.newLine();
			output.flush();
			output.write("  12 Any Happiness/Surprise");
			output.newLine();
			output.flush();
			output.write("  22 Any positive");
			output.newLine();
			output.flush();
			output.write("  48 Any negative");
			output.newLine();
			output.flush();
			output.write("  _99 Overall emotional intensity");
			output.newLine();
			output.flush();
			output.write("  _98 Error Variance");
			output.newLine();
			output.flush();
			output.write("  _97 Standard Error");
			output.newLine();
			output.flush();
			output.write("S_FT1 &#0025; SUMMARY:");
			output.newLine();
			output.flush();
			output.write("FT1 <question_name>. WHICH OF THESE FACES BEST EXPRESSES HOW YOU THINK PEOPLE WOULD FEEL ABOUT THIS <specify>?");
			output.newLine();
			output.flush();
			output.write("EMOTION_T FACE TRACE");
			output.newLine();
			output.flush();
			output.write("EMOT FACE TRACE");
			output.newLine();
			output.flush();
			output.write("  1 Surprise");
			output.newLine();
			output.flush();
			output.write("  2 Happy");
			output.newLine();
			output.flush();
			output.write("  3 Neutral");
			output.newLine();
			output.flush();
			output.write("  4 Sadness");
			output.newLine();
			output.flush();
			output.write("  5 Fear");
			output.newLine();
			output.flush();
			output.write("  6 Anger");
			output.newLine();
			output.flush();
			output.write("  7 Disgust");
			output.newLine();
			output.flush();
			output.write("  8 Contempt");
			output.newLine();
			output.flush();
			output.write("  12 Any Happiness/Surprise");
			output.newLine();
			output.flush();
			output.write("  22 Any positive");
			output.newLine();
			output.flush();
			output.write("  48 Any negative");
			output.newLine();
			output.flush();
			output.write("  _99 Overall emotional intensity");
			output.newLine();
			output.flush();
			output.write("  _98 Error Variance");
			output.newLine();
			output.flush();
			output.write("  _97 Standard Error");
			output.newLine();
			output.flush();
			output.write("S_FT2 &#0025; SUMMARY:");
			output.newLine();
			output.flush();
			output.write("FT2_1 <question_name2>(1). DEGREE TO WHICH PEOPLE WOULD FEEL EMOTION ABOUT THIS <specify>: Surprise");
			output.newLine();
			output.flush();
			output.write("FT2_2 <question_name2>(2). DEGREE TO WHICH PEOPLE WOULD FEEL EMOTION ABOUT THIS <specify>: Happiness");
			output.newLine();
			output.flush();
			output.write("FT2_4 <question_name2>(4). DEGREE TO WHICH PEOPLE WOULD FEEL EMOTION ABOUT THIS <specify>: Sadness");
			output.newLine();
			output.flush();
			output.write("FT2_5 <question_name2>(5). DEGREE TO WHICH PEOPLE WOULD FEEL EMOTION ABOUT THIS <specify>: Fear");
			output.newLine();
			output.flush();
			output.write("FT2_6 <question_name2>(6). DEGREE TO WHICH PEOPLE WOULD FEEL EMOTION ABOUT THIS <specify>: Anger");
			output.newLine();
			output.flush();
			output.write("FT2_7 <question_name2>(7). DEGREE TO WHICH PEOPLE WOULD FEEL EMOTION ABOUT THIS <specify>: Disgust");
			output.newLine();
			output.flush();
			output.write("FT2_8 <question_name2>(8). DEGREE TO WHICH PEOPLE WOULD FEEL EMOTION ABOUT THIS <specify>: Contempt");
			output.newLine();
			output.flush();
			output.write("  1 A little (1.)");
			output.newLine();
			output.flush();
			output.write("  2 Moderately (2.)");
			output.newLine();
			output.flush();
			output.write("  3 A lot (3.)");
			output.newLine();
			output.flush();
			output.write("  _99 Mean");
			output.newLine();
			output.flush();
			output.write("  _98 Error Variance");
			output.newLine();
			output.flush();
			output.write("  _97 Standard Error");
			output.newLine();
			output.flush();
			output.newLine();
			output.flush();
			output.newLine();
			output.flush();

		} else if (question.getType() != null && !question.getType().isEmpty()) {
			output.write(question.getTitle().toUpperCase() + " "
					+ question.getTitle().toUpperCase() + ". "
					+ question.getLabel().toUpperCase());
			output.newLine();
			output.flush();
			int i = 1;
			for (Answer answer : question.getAnswers()) {
				output.write("   " + i + " " + answer.getLabel());
				output.newLine();
				output.flush();
				i++;
			}
			if (question.getType().equals("multi")) {
				output.write("   _99 Mean");
				output.newLine();
				output.flush();
				output.write("   _98 Error Variance");
				output.newLine();
				output.flush();
				output.write("   _97 Standard Error");
				output.newLine();
				output.flush();
			}
			output.newLine();
			output.flush();
			output.newLine();
			output.flush();
		}

	}

	public FileInputStream getInputFile() {
		return inputFile;
	}

	public void setInputFile(FileInputStream inputFile) {
		this.inputFile = inputFile;
	}

	private static String getPrefix(Cell input) {
		if (input == null) {
			return "";
		}
		String cellValue = input.getStringCellValue();
		if (cellValue == null || cellValue.isEmpty()) {
			return "";
		}
		return cellValue.split("_")[0];
	}

	private static boolean isQuestionOfType(String questionType,
			Question question) {
		return questionType != null && question != null
				&& questionType.equals(question.getType());
	}

	private void logInfo(String message) {
		DateFormat dateFormat = new SimpleDateFormat("HH:mm:ss ");
		logBuffer.append(dateFormat.format(new Date()));
		logBuffer.append(message);
		logBuffer.append(System.getProperty("line.separator"));
	}

	public String getLog() {
		return logBuffer.toString();
	}

}
