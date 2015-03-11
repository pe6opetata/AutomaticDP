package classes;

/**
 * The class contains the label and the code of a certen answer.
 * 
 * @author Peter Tzokov
 * 
 * 
 */
public class Answer extends Question {

	private String label;
	private int code;

	/**
	 * Default constructor.
	 */
	public Answer() {
		super();
		label = null;
		code = 0;
	}

	/**
	 * Constructor with parameters.
	 * 
	 * @param label
	 *            the label of the answer.
	 * @param code
	 *            the code of the answer.
	 */
	public Answer(String label, int code) {
		super();
		this.label = label;
		this.code = code;
	}

	/**
	 * Returns the label of the answer.
	 */
	public String getLabel() {
		return label;
	}

	/**
	 * Sets a label.
	 * 
	 * @param label
	 *            The label you want to set.
	 */
	public void setLabel(String label) {
		this.label = label;
	}

	/**
	 * 
	 * @return Returns the code of the answer.
	 */
	public int getCode() {
		return code;
	}

	/**
	 * 
	 * @param code
	 *            Sets the code of the answer.
	 */
	public void setCode(int code) {
		this.code = code;
	}

}
