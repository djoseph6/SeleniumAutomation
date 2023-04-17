import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.Assert;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

public class RegisterTest {
	
	WebDriver chrDrive;

//	public static void main(String[] args) {
//		
//		
//		WebDriver chrDrive = new GoogleChromeWebProperties().drive;
//		
//		try {
//			visitGoogleWebsite(chrDrive);
//		}catch(Exception e) {
//			e.getMessage();
//		}finally {
//			chrDrive.quit();
//		}
//		
//		
//		
//	}
	
	@BeforeTest 
	void setupTestCase() {
		chrDrive = new GoogleChromeWebProperties().drive;
	}
	
	@Test
	void visitGoogleWebsite() {
		chrDrive = new ChromeDriver();
		chrDrive.get("https://www.google.com");
		String webTitle = chrDrive.getTitle();
		System.out.println(webTitle);
		Assert.assertEquals(webTitle, "Google");
		
	}

}
