package excelOperate;

public class RowContent {

	private int volumeNum;
	private int issueNum;
	private int quotedNum;
	private String paperTitle;
	private String author;
	private String email;

	public RowContent() {
	}

	public int getVolumeNum() {
		return volumeNum;
	}

	public void setVolumeNum(int volumeNum) {
		this.volumeNum = volumeNum;
	}

	public int getIssueNum() {
		return issueNum;
	}

	public void setIssueNum(int issueNum) {
		this.issueNum = issueNum;
	}

	public int getQuotedNum() {
		return quotedNum;
	}

	public void setQuotedNum(int quotedNum) {
		this.quotedNum = quotedNum;
	}

	public String getPaperTitle() {
		return paperTitle;
	}

	public void setPaperTitle(String paperTile) {
		this.paperTitle = paperTile;
	}

	public String getAuthor() {
		return author;
	}

	public void setAuthor(String author) {
		this.author = author;
	}

	public String getEmail() {
		return email;
	}

	public void setEmail(String email) {
		this.email = email;
	}

	@Override
	public String toString() {
		return "PaperTitle: " + getPaperTitle() + "\n" + "Author: " + getAuthor() + "\n" + "Volume: " + getVolumeNum()
				+ "; " + "Issue: " + getIssueNum() + "; " + "QuotedNum: " + getQuotedNum() + "; " + "Email: "
				+ getEmail() + "\n";
	}
}
