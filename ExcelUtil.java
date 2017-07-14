package st.excel.util;

public class ExcelUtil {
	/**
	 * 开始行
	 */
	private int row_start;

	/**
	 * 结束行
	 */
	private int row_end;

	/**
	 * 开始列
	 */
	private int col_start;

	/**
	 * 结束列
	 */
	private int col_end;

	/**
	 * 单转换区域
	 * 
	 * @param area
	 *            区域 C,15-B,16
	 */
	public void dealArea(String area) {
		/* 拆分区域 */
		String[] ap = area.split("-");
		String _start = ap[0].trim();
		String _end = ap[1].trim();
		String[] temp = _start.split(",");
		String _x = temp[0];
		String _y = temp[1];

		this.row_start = Integer.parseInt(_y) - 1;
		this.col_start = CovertColToNum(_x);
		temp = _end.split(",");
		_x = temp[0];
		_y = temp[1];
		if (_y.equals("*")) {
			//不定行
			this.row_end = -1;
		} else {
			this.row_end = Integer.parseInt(_y) - 1;
		}
		
		if (_x.equals("*")) {
			//不定列
			this.col_end = -1;
		} else {
			this.col_end = CovertColToNum(_x);
		}
	}

	/**
	 * 只支持两位字母
	 * 
	 * @param letter
	 *            字母列号
	 * @return 列号
	 */
	public int CovertColToNum(String letter) {
		if (letter.length() == 1) {
			return letter.charAt(0) - 65;
		} else if (letter.length() == 2) {
			int l = letter.charAt(0) - 64;
			int p = letter.charAt(1) - 65;
			return l * 26 + p;
		}
		return 0;
	}

	public int getRow_start() {
		return row_start;
	}

	public int getRow_end() {
		return row_end;
	}

	public int getCol_start() {
		return col_start;
	}
 
	public int getCol_end() {
		return col_end;
	}

	// 46
	public static void main(String[] args) {
		ExcelUtil e = new ExcelUtil();
		int p = e.CovertColToNum("AW");// ;-e.CovertColToNum("D");
		System.out.print(p);
	}
}
