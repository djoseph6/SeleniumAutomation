import org.openqa.selenium.WebDriver;

public class GoogleChromeWebProperties {
	
	private WebDriver drive;
	
	
	GoogleChromeWebProperties(){
		System.setProperty("selenium-chrome-driver-4.8.1", "C:\\Program Files\\Selenium");
		
	}
	
	public WebDriver getWebDriver() {
		return drive;
		
	}
	

}
