import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Properties;
import java.util.concurrent.TimeUnit;

import javax.activation.DataHandler;
import javax.activation.FileDataSource;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxBinary;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.firefox.FirefoxProfile;
import org.openqa.selenium.firefox.internal.ProfilesIni;
import org.openqa.selenium.support.ui.Select;
import org.testng.annotations.Test;

import com.opencsv.CSVReader;

public class Mantis {

	static String xlsxReport ;
	public static void main(String[] args) throws AWTException, InterruptedException 
	{		
		
		//ProfilesIni profile = new ProfilesIni();
		//FirefoxProfile myProfile = profile.getProfile("profileApexLink");
		//WebDriver driver = new FirefoxDriver(myProfile);
		//WebDriver driver = new FirefoxDriver(new FirefoxBinary(new File("C:\\Program Files (x86)\\Mozilla Firefox\\firefox.exe")), myProfile);
		
		System.setProperty("webdriver.chrome.driver", System.getProperty("user.dir") + "\\Driver\\chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		driver.manage().timeouts().pageLoadTimeout(10, TimeUnit.SECONDS);		
		
		driver.get("http://mantis.grotal.com/login_page.php");
		//Fill Username & Password, Click Login
		driver.findElement(By.name("username")).sendKeys("yashu");
		driver.findElement(By.name("password")).sendKeys("yashmit@123");
		driver.findElement(By.className("button")).click();
		
		driver.findElement(By.linkText("View Issues")).click();
		System.out.println("Clicked on view issues");
		
		
		Select sel = new Select(driver.findElement(By.name("project_id")));
		sel.selectByVisibleText("All Projects");//Project tp Select from DDL: Oujra, iFag
		driver.findElement(By.xpath("//input[@value='Switch']")).click();
		System.out.println("selected all projects");
		
		
		driver.findElement(By.name("reset_query_button")).click();//Reset Filters
		System.out.println("Reset query button");
		
		
		driver.findElement(By.id("hide_status_filter")).click();//Hide status as None
		sel = new Select(driver.findElement(By.xpath("//td[@id='hide_status_filter_target']/select")));
		sel.selectByVisibleText("[none]");
		System.out.println("selected none hide status");
		//Use date filter
		driver.findElement(By.id("do_filter_by_date_filter")).click();
		driver.findElement(By.name("do_filter_by_date")).click();
		
		String []startEndDates = dateFilterStrings();
		//Start Date
		Integer month = Integer.parseInt(startEndDates[0].split("/")[1]);
		Integer day = Integer.parseInt(startEndDates[0].split("/")[2]);
		sel = new Select(driver.findElement(By.name("start_year")));
		sel.selectByVisibleText(startEndDates[0].split("/")[0]);
		sel = new Select(driver.findElement(By.name("start_month")));
		sel.selectByValue(month.toString());	
		sel = new Select(driver.findElement(By.name("start_day")));
		sel.selectByVisibleText(day.toString());
		
		//End Date
		month = Integer.parseInt(startEndDates[1].split("/")[1]);
		day = Integer.parseInt(startEndDates[1].split("/")[2]);
		sel = new Select(driver.findElement(By.name("end_year")));
		sel.selectByVisibleText(startEndDates[1].split("/")[0]);
		sel = new Select(driver.findElement(By.name("end_month")));
		sel.selectByValue(month.toString());	
		sel = new Select(driver.findElement(By.name("end_day")));
		sel.selectByVisibleText(day.toString());
		
		driver.findElement(By.id("custom_field_8_filter")).click();//Type Select as Code review
		sel = new Select(driver.findElement(By.xpath("//td[@id='custom_field_8_filter_target']/select")));
		sel.selectByVisibleText("Code Review");
		//Thread.sleep(5000);
		System.out.println("selected code review");
		driver.findElement(By.
				xpath("//div[@id='filter_open']//input[@name='filter']")).click();
		
		driver.findElement(By.linkText("CSV Export")).click();
		
		/*Robot robo= new Robot();
		robo.keyPress(KeyEvent.VK_DOWN);
		robo.keyPress(KeyEvent.VK_DOWN);
		//robo.keyPress(KeyEvent.VK_ACCEPT);
		robo.keyRelease(KeyEvent.VK_ENTER);*/
		
	     csvToXlsx();
		driver.quit();
		
		String mailText = "Please find the attached 'Code Review Report' for previous week."
				+ System.lineSeparator() + ""
				+ System.lineSeparator() + ""
				+ System.lineSeparator() + ""
				+ System.lineSeparator() + "Regards"
				+ System.lineSeparator() + "Seasia Infotech- A Cmmi Level 5 Company"
				+ System.lineSeparator() + "Yashu Kapila| Senior Manager"
				+ System.lineSeparator() + "Skype: Yashu.kapila"
				+ System.lineSeparator() + "Contact: +91.172.5218500 (O) Ext: 72009 | +91.977.900.0714 (M)"
				+ System.lineSeparator() + "US Number : +1.240.241.6894"
				+ System.lineSeparator() + "www.seasiainfotech.com";
		String []to = {"kaurtaranbir@SEASIAINFOTECH.COM"};//{"BrarArshdeep@SEASIAINFOTECH.COM","VermaGanesh@SEASIAINFOTECH.COM"};
		String []cc =  {"kaurtaranbir@SEASIAINFOTECH.COM"}; // {"yashukapila@seasiainfotech.com"};//{"deepak@SEASIAINFOTECH.COM"};//
		String []bcc = {"kaurtaranbir@SEASIAINFOTECH.COM"};
		
		sendMail(to, cc, bcc, "yashukapila@seasiainfotech.com",
				"Code Review Report", mailText);	//	
	}
	
	/**
	 * 
	 * @return yyyy/MM/dd format
	 */
	
	public static String[] dateFilterStrings()
	{
		/*Date lastMondayDate = new Date(System.currentTimeMillis()-24*60*60*1000*(7));
		
		System.out.println("lastMondayDate::"+lastMondayDate);
		
		Date yesterdayDate = new Date(System.currentTimeMillis()-24*60*60*1000*(1));
		System.out.println("yesterdayDate::"+yesterdayDate);*/
		String []returnDates = new String[2];
		
		returnDates[0] = "2019/07/01";          //new SimpleDateFormat("yyyy/MM/dd").format(lastMondayDate);
		
		System.out.println("returnDates[0]::"+returnDates[0]);
		returnDates[1] =   "2019/07/08";                //new SimpleDateFormat("yyyy/MM/dd").format(yesterdayDate);
		System.out.println("returnDates[1]::"+returnDates[1]);
		
		/*lastMondayDate::Mon Jul 01 12:34:25 IST 2019
		yesterdayDate::Sun Jul 07 12:34:25 IST 2019
		returnDates[0]::2019/07/01
		returnDates[1]::2019/07/07*/
		
		return returnDates;
	}
	
	public static void csvToXlsx(){
		try {
			Thread.sleep(3000);
	    	String filepath = getLatestFileFromDir("C:\\Users\\kaurtaranbir\\Downloads");
	        CSVReader reader = new CSVReader(new FileReader(filepath));	        
	        
	        System.out.println("Reader::"+reader);
	        String[] line;
	        xlsxReport = "D:\\Taran\\mantis" +"\\CodeReviewReport.xlsx";  
	        FileInputStream fis = new FileInputStream(xlsxReport);
			XSSFWorkbook workBook = new XSSFWorkbook(fis);
			int index = workBook.getSheetIndex("yashu");
			XSSFSheet sheet = workBook.getSheetAt(index);
			int lastRow = sheet.getLastRowNum();
			
			for (int j = 1; j<=lastRow; j++)
			{
				XSSFRow row = sheet.getRow(j);
        		sheet.removeRow(row);
        		System.out.println("j::"+j);
			}
			//reader.
			int i = 0;
			for( ; (line = reader.readNext()) != null; i++){
				
				XSSFRow row = sheet.createRow(i);
				for(int j=0; j<line.length; j++){
	            	XSSFCell cell = row.createCell(j);	            		            	
	            	cell.setCellValue(line[j]);
	            }        
			}	        
			
			FileOutputStream fileOutputStream =  
					new FileOutputStream(xlsxReport);
			
			System.out.println("fileOutputStream::"+fileOutputStream);
	        workBook.write(fileOutputStream);
	        fileOutputStream.close();
	        workBook.close();
	        reader.close();
	        System.out.println("Done");    
	        
	    } catch (Exception ex) {
	    	ex.printStackTrace();
	        System.out.println(ex.getMessage()+"Exception in try");
	    }
	}
	
	public static String getLatestFileFromDir(String dirPath) {
		File dir = new File(dirPath);
		File[] files = dir.listFiles();
		if (files == null || files.length == 0) {
			return null;
		}

		File lastModifiedFile = files[0];
		for (int i = 1; i < files.length; i++) {
			if (lastModifiedFile.lastModified() < files[i].lastModified()) {
				lastModifiedFile = files[i];
			}
		}
		
		return lastModifiedFile.getAbsolutePath();
	}
	
	public static void sendMail(String[] to, String[] cc, String[] bcc, String from, String subject, String msg)
	{
	      String host = "webmail.seasiainfotech.com"; //"smtp.seasiainfotech.com";
	      Properties properties = System.getProperties();
	      properties.setProperty("mail.smtp.host", host);	      
	      Session session = Session.getDefaultInstance(properties);

	      try {
	    	  Thread.sleep(1000);
	    	  //FileInputStream fis = new FileInputStream("C:\\Users\\singhjasmeet\\Desktop\\Test.xlsx");
	    	  File attachmentSource = new File(xlsxReport);
				//XSSFWorkbook workBook = new XSSFWorkbook(fis);
	    	  MimeBodyPart attachment = new MimeBodyPart();
	    	  attachment.setDataHandler(new DataHandler(new FileDataSource(attachmentSource)));
	    	  attachment.setFileName(attachmentSource.getName());
	    	  MimeBodyPart text = new MimeBodyPart(); 
	    	  text.setText(msg);
	    	  Multipart mp = new MimeMultipart();
	    	  mp.addBodyPart(attachment);
	    	  mp.addBodyPart(text);
	         MimeMessage message = new MimeMessage(session);
	         message.setFrom(new InternetAddress(from));
	         
	         for (String string : to) {
	        	 message.addRecipient(Message.RecipientType.TO, new InternetAddress(string));
			}
	         for (String string : cc) {
	        	 message.addRecipient(Message.RecipientType.CC, new InternetAddress(string));
			}
	         for (String string : bcc) {
	        	 message.addRecipient(Message.RecipientType.BCC, new InternetAddress(string));
			}
	         /*message.addRecipient(Message.RecipientType.TO, new InternetAddress(to));
	         message.addRecipient(Message.RecipientType.TO, new InternetAddress(to));
	         message.addRecipient(Message.RecipientType.CC, new InternetAddress(to));*/
	         message.setSubject(subject);
	         //message.setText(msg);
	         message.setContent(mp);
	         message.saveChanges();

	         
	         
	         Transport.send(message);
	         System.out.println("Sent message successfully....");
	      }catch (MessagingException | InterruptedException ex) {
	         ex.printStackTrace();
	      }		
	}

}
