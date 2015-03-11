package classes;

import java.util.ArrayList;
import java.util.List;

public class Question {

	private String ID;
	private String title;
	private String label;
	private String type;
	private int startCol;
	private int endCol;
	private List<Answer> answers;

	public Question() {
		this.title = null;
		this.label = null;
		this.type = null;
		this.startCol = 0;
		this.answers = new ArrayList<Answer>();
	}

	public String getID() {
		return ID;
	}

	public void setID(String iD) {
		ID = iD;
	}

	public String getTitle() {
		return title;
	}

	public void setTitle(String title) {
		this.title = title;
	}

	public String getLabel() {
		return label;
	}

	public void setLabel(String label) {
		this.label = label;
	}

	public String getType() {
		return type;
	}

	public void setType(String type) {
		this.type = type;
	}

	public int getStartCol() {
		return startCol;
	}

	public void setStartCol(int startCol) {
		this.startCol = startCol;
	}

	public int getEndCol() {
		return endCol;
	}

	public void setEndCol(int endCol) {
		this.endCol = endCol;
	}

	public List<Answer> getAnswers() {
		return answers;
	}

	public void setAnswers(List<Answer> answers) {
		this.answers = answers;
	}

	public void addAnswer(Answer answer) {
		answers.add(answer);
	}
}
