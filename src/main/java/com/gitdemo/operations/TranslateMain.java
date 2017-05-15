package com.gitdemo.operations;

import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.security.GeneralSecurityException;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Date;
import java.util.Enumeration;
import java.util.List;
import java.util.Properties;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.gitdemo.trans.DemoTranslate;
import com.gitdemo.util.CommonUtil;

public class TranslateMain {
	
	private static FileDataStoreFactory DATA_STORE_FACTORY;
	private static HttpTransport HTTP_TRANSPORT;
	private static final JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();
	private static String inputFileName = "";
	private static String outputFileName = "";
	private static int batchSize = 0;
	
	private static final List<String> SCOPES = Arrays.asList(SheetsScopes.DRIVE, SheetsScopes.DRIVE_FILE,
			SheetsScopes.SPREADSHEETS);

	static {
		try {
			HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
			DATA_STORE_FACTORY = new FileDataStoreFactory(CommonUtil.DATA_STORE_DIR);
		} catch (Throwable t) {
			t.printStackTrace();
			System.exit(1);
		}
	}
	public static void readProperties(){
		InputStream input = null;
		Properties prop = new Properties();
		input = TranslateMain.class.getClassLoader().getResourceAsStream("config.properties");
		try {
			prop.load(input);
			Enumeration<?> e = prop.propertyNames();
			while (e.hasMoreElements()) {
				String key = (String) e.nextElement();
				
				if(key.equals("batchsize")){
					batchSize = Integer.parseInt(prop.getProperty(key));					
				}else if(key.equals("inputfile")){					
					inputFileName = prop.getProperty(key);					
				}else{
					outputFileName = prop.getProperty(key);					
				}
				
			}
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		
	}
	private static Credential authorize() throws IOException {

		
		// Load client secrets.
		InputStream in = DemoTranslate.class.getClassLoader().getResourceAsStream("client_secret.json");
		GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));
		// Build flow and trigger user authorization request.
		GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(HTTP_TRANSPORT, JSON_FACTORY,
				clientSecrets, SCOPES).setDataStoreFactory(DATA_STORE_FACTORY)
						.build();
		Credential credential = new AuthorizationCodeInstalledApp(flow, new LocalServerReceiver()).authorize("user");
		System.out.println("INFO:: Credentials saved to :::::::::::::::: " + CommonUtil.DATA_STORE_DIR.getAbsolutePath());
		return credential;
	}
	
	private static Sheets createSheetsService() throws IOException, GeneralSecurityException {

		Credential credential = authorize();
		System.out.println("INFO:: Sheets Service creation completed:::::::::::::::::");
		return new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, credential).setApplicationName(CommonUtil.APPLICATION_NAME)
				.build();
		
	}


	public static void main(String[] args) throws Exception{
		
		System.out.println("::::::::::::::::::::::::START::::::::::::::::::::::::");
		Date startTime = (Date)Calendar.getInstance().getTime();
		Sheets sheetsService = null;
		System.out.println("INFO:: Read configuration  < =====< =====< =====");
		TranslateMain.readProperties();
		System.out.println("INFO:: Read source data and Writing into CSV file  =====> =====> =====>");
		DataMoveOperation.readXLSXFile(inputFileName,outputFileName);
		
		try {
			
			sheetsService = createSheetsService();
			DemoTranslate.sheetOperations(sheetsService,outputFileName,batchSize);
			
			Date endTime = (Date)Calendar.getInstance().getTime();
			System.out.println("Info:: "+"Translate Process Completed in "+(endTime.getTime()-startTime.getTime())/1000+" Seconds");
		} catch (IOException e) {
			e.printStackTrace();
		} catch (GeneralSecurityException e) {
			e.printStackTrace();
		}
		

	}

}
