package cn.javaex.officejj.common.entity;

public class RGB {
	private int red;
	private int green;
	private int blue;
	
	public RGB() {
		
	}
	
	public RGB(String colorStr) {
		if (colorStr.startsWith("#")) {
			colorStr = colorStr.substring(1);
		}
		this.red = Integer.valueOf(colorStr.substring(0, 2), 16 );
		this.green = Integer.valueOf(colorStr.substring(2, 4), 16 );
		this.blue = Integer.valueOf(colorStr.substring(4, 6), 16 );
		
	}
	
	public int getRed() {
		return red;
	}
	public void setRed(int red) {
		this.red = red;
	}
	public int getGreen() {
		return green;
	}
	public void setGreen(int green) {
		this.green = green;
	}
	public int getBlue() {
		return blue;
	}
	public void setBlue(int blue) {
		this.blue = blue;
	}
	
}
