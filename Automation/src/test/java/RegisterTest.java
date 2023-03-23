import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class RegisterTest {

	public static void main(String[] args) {
		
		GoogleChromeWebDriver driver;
		WebDriver chrDrive = null;
		
		try {
			visitGoogleWebsite(chrDrive);
		}catch(Exception e) {
			e.getMessage();
		}finally {
			chrDrive.quit();
		}
		
		
		
	}
	
	static void visitGoogleWebsite(WebDriver drive) {
		drive = new ChromeDriver();
		drive.get("https://www.google.com");
	}

}
