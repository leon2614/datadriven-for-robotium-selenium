package com.dfrs;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.HttpURLConnection;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import android.app.Activity;
import android.content.ComponentName;
import android.content.Context;
import android.content.Intent;
import android.content.ServiceConnection;
import android.os.Environment;
import android.os.IBinder;
import android.os.RemoteException;
import android.util.Log;

import com.jayway.android.robotium.solo.Solo;

public class KoolJ_datadriven {
	
	String test_xls;
	String suite_xls;
	String map_xls;
	String batch_xls;
	String config_xls;
	String KOOLJ_log;
	String[][] outputReport;
	Object[][] data_suite;
	Object[][] data_test;
	Object[][] data_key;
	Object[][] data_url_batch;
	int read_first = 1;
	int file_download_done = 0;
	
	
	
	//Open CONFIG to BATCH,SUITE,TEST files
	public void openconfig(String config_xls, String output_xls, Solo solo){
		Object[][] data_batch = CreateDataFromCSV(config_xls);

		//check NULL data_batch
		if (data_batch == null) {
			Log.e("KOOLJ_log", "DATA IS NULL");
			KOOLJ_log=KOOLJ_log+"\n"+"DATA IS NULL";
		}
		else
		{
			Log.e("KOOLJ_log", "DATA IS AVAIL");
			KOOLJ_log=KOOLJ_log+"\n"+"DATA IS AVAIL";
			
			//if files from HTTP, download them
			file_download_done = 1;
			data_url_batch = CreateDataFromCSV("/url_batch.xls");
			for (int i_d=1; i_d< data_url_batch.length; i_d++)
			{					
				if (data_url_batch[0][1].toString().equals("yes"))
				{
					URLfile(data_url_batch[i_d][1].toString(),data_url_batch[i_d][0].toString());
				}
				else 
				{
					file_download_done = 3;
				}	
				
			}
				
			//Find to run BATCH	
			for (int i=0; i< data_batch.length; i++)
			{	
				if (file_download_done > 1)
				{
					//When downloading done, read files
					if (data_batch[i][0].equals("xls_batch_url")) 
					{
						//Find to run SUITE
						String data_suite_var="/" + data_batch[i][1].toString() + ".xls";
						KOOLJ_log=KOOLJ_log+"\n"+"RUN BATCH: "+ data_suite_var;
						Log.e("KOOLJ_BATCH: ", data_suite_var);
						data_suite = CreateDataFromCSV(data_suite_var);
						break;
					}
				}
			}
			
			//Find to run TEST
			if (file_download_done > 2)
			{
				for (int ii=0; ii< data_suite.length; ii++)
				{
					
					String data_test_var="/" + data_suite[ii][0].toString() + ".xls";
					KOOLJ_log=KOOLJ_log+"\n"+"RUN SUITE:______ "+ data_test_var;
					Log.e("KOOLJ_SUITE_"+ii+": ", data_test_var);
					data_test = CreateDataFromCSV(data_test_var);
					
					//Find to run KEY
					for (int iii=0; iii< data_test.length; iii++)
					{
						String data_key_var="/" + data_test[iii][0].toString() + ".xls";
						KOOLJ_log=KOOLJ_log+"\n"+"RUN TEST:______ "+ data_key_var;
						Log.e("KOOLJ_TEST_"+iii+": ", data_key_var);
						data_key = CreateDataFromCSV(data_key_var);
						
						//Run each KEY
						String[] keyx_label=new String[data_key.length];
						String[] valuex_label=new String[data_key.length];
						String[] key_for=new String[data_key.length];
						String[] valuestart_for=new String[data_key.length];
						String[] valueend_for=new String[data_key.length];
						String[] valueacce_for=new String[data_key.length];
						String[] key_endfor=new String[data_key.length];


						int for_step = 0;
						int for_step_backward = 0;						
						int endfor_step = 0;
						int keyx_label_step = 0;
						int iiii_label = 0;
						int value_valuestart_for = 0;
						int value_valueend_for = 0;
						int value_valueacce_for = 0;
						int varstore_kv_step = 0;
						int varstore_step = 0;
						int varstore_count = 0;
						
						//Store LABEL if have
						for (int iiii=iiii_label; iiii< data_key.length; iiii++)
						{
							String key_target = data_key[iiii][1].toString();
							if (key_target.equals("label"))
							{
								
								keyx_label[keyx_label_step] = ""+iiii;
								valuex_label[keyx_label_step] = data_key[iiii][2].toString();
								
								Log.e("KOOLJ_label_", data_key[iiii][2].toString());
								keyx_label_step++;
							}
						}
						//Count FOR..ENDFOR if have
						for (int iiii=iiii_label; iiii< data_key.length; iiii++)
						{
							String key_target = data_key[iiii][1].toString();
							if (key_target.equals("for"))
							{
								key_for[for_step] = ""+iiii;
								for_step++;
								for_step_backward = for_step;
							}
							if (key_target.equals("endfor"))
							{
								for_step_backward--;
								key_endfor[endfor_step] = key_for[for_step_backward].toString();
								endfor_step++;
								
							}
						}
						
						//Count STORE
						for (int iiii=iiii_label; iiii< data_key.length; iiii++)
						{
							String key_target = data_key[iiii][1].toString();
							if (key_target.equals("storevar"))
							{
								varstore_count++;
							}
						}
						
						//Store values of STORE if have
						Object[][] varstore_kv=new Object[varstore_count][2];
						for (int iiii=iiii_label; iiii< data_key.length; iiii++)
						{
							String key_target = data_key[iiii][1].toString();
							if (key_target.equals("storevar"))
							{
								varstore_kv[varstore_step][0] = data_key[iiii][2].toString();
								varstore_kv[varstore_step][1] = data_key[iiii][3].toString();
								varstore_step++;									
							}
						}
						
						//Search to change the STORE VAR if it is repeated
						/*
						for (int iiii=iiii_label; iiii< data_key.length; iiii++)
						{
							String key_target = data_key[iiii][1].toString();
							if (key_target.equals("storevar"))
							{
								String var_temp = data_key[iiii][2].toString();
								String var_temp3 = data_key[iiii][3].toString();
								for (int ix = 0; ix< varstore_kv.length; ix++)
								{
									String var_temp2 = varstore_kv[ix][0].toString();
									if (var_temp2.equals(var_temp))
									{
										for (int iz = 0; iz< varstore_kv.length; iz++)
										{
											if (!var_temp2.equals(var_temp3))
											{
												varstore_kv[ix][1] = data_key[iiii][3].toString();
												break;
											}	
										}	
									}
								}
							}
						}
						*/
						
						//Count IF..ENDIF if have
						for (int iiii=iiii_label; iiii< data_key.length; iiii++)
						{

						}
						//Run each KEY
						for (int iiii=iiii_label; iiii< data_key.length; iiii++)
						{
							
							String key_target = data_key[iiii][1].toString();
							if(key_target.equals("sleep"))
							{
								int key_value = Integer.parseInt(data_key[iiii][2].toString());
								
								solo_sleep(key_value, solo);
							}
							else if(key_target.equals("for"))
							{
								//for (int i = 0; i< varstore_kv.length; i++)
								//{
									//duyet varstore_kv
										//so sanh [i][0] voi data_key[iiii][3].toString()
											// lay ra varstore_kv[i][1]
									//duyet varstore_kv
										//so sanh [i][0] voi data_key[iiii][4].toString()	
											// lay ra varstore_kv[i][1]
									//duyet varstore_kv
										//so sanh [i][0] voi data_key[iiii][5].toString()
											//// lay ra varstore_kv[i][1]
											
								//}
							}
							else if(key_target.equals("endfor"))
							{
								/*
								if ( value_valuestart_for <= value_valueend_for)
								{
									value_valuestart_for = value_valuestart_for + value_valueacce_for;
									for (int i = 0; i< valueend_for_yes.length; i++)
									{
										int key_valueend_for_yes = Integer.parseInt(valueend_for_yes[i].toString());
										if (key_valueend_for_yes == iiii)
										{
											iiii_label = Integer.parseInt(valueend_for_yes[i].toString());
											iiii = iiii_label;
											break;
										}
									}
								}
								*/
							}
							else if(key_target.equals("storevar"))
							{
								//Search to change the VAR
								/*
								Log.e("KOOLJ_STEP" , "_________");
								for (int iz = 0; iz< varstore_kv.length; iz++)
								{
									Log.e("KOOLJ_flow " , ""+iz);
									String var_temp = varstore_kv[iz][0].toString();
									if (var_temp.equals(data_key[iiii][2].toString()))
									{
										String var_temp3 = " ";
										for (int ix = 0; ix< varstore_kv.length; ix++)
										{
											String var_temp2 = varstore_kv[ix][0].toString();
											Log.e("KOOLJ_temp2 " , var_temp2);
											Log.e("KOOLJ_datak3 " , data_key[iiii][3].toString());
											if (var_temp2.equals(data_key[iiii][3].toString()))
											{
												var_temp3 = varstore_kv[ix][1].toString();
												//varstore_kv[iz][1] = varstore_kv[ix][1].toString();
												Log.e("KOOLJ_eq" + "_" + iz + " " + varstore_kv[iz][0].toString(), varstore_kv[iz][1].toString());
												//break;
											}
											
										}
										varstore_kv[iz][1] = var_temp3;
									//break;
									}
									
								}
								*/
							}	
							else if(key_target.equals("tracevar"))
							{
								/*
								for (int i = 0; i< varstore_kv.length; i++)
								{
									String var_temp = varstore_kv[i][0].toString();
									if (var_temp.equals(data_key[iiii][2].toString()))
									{									
										Log.e("KOOLJ_TRACE" + i + "_"+varstore_kv[i][0].toString(), varstore_kv[i][1].toString());
										break;
									}
								}
								*/
							}							
							else if(key_target.equals("getCurrentActivity"))
							{
								solo_assertCurrentActivity ("View current activity fail!", solo.getCurrentActivity().getClass(), solo);
								
							}
							else if(key_target.equals("screenshot"))
							{
								//Float key_value1 = Float.valueOf(data_key[iiii][2].toString());
								//Float key_value2 = Float.valueOf(data_key[iiii][3].toString());
								//asl_screenshot(key_value1, key_value2, solo);
								
							}
							else if(key_target.equals("sendKey"))
							{
								int key_value = Integer.parseInt(data_key[iiii][2].toString());
								
								solo_key(key_value, solo);
								
							}
							else if (key_target.equals("searchText"))
							{
								
								solo_searchtext(data_key[iiii][2].toString(), solo);
								
							}
							else if (key_target.equals("goBack"))
							{
								
								solo_back(solo);
								
								
							}
							else if (key_target.equals("enterText"))
							{
								
								solo_enterkey(Integer.parseInt(data_key[iiii][2].toString()), data_key[iiii][3].toString(), solo);
							}
							else if (key_target.equals("clickOnButton"))
							{
								
								solo_clickonbutton(data_key[iiii][2].toString(), solo);
								
							}
							
							else if (key_target.equals("goto"))
							{
								Log.e("KOOLJ_goto_", data_key[iiii][2].toString());
								int i_labelhave = 0;
								for (int i_label = 0; i_label< keyx_label_step; i_label++)
								{
									String valuex_target = valuex_label[i_label].toString();
									if (valuex_target.equals(data_key[iiii][2].toString()))
									{
										i_labelhave = 1;
										iiii_label = Integer.parseInt(keyx_label[i_label]);	
										iiii = iiii_label;										
										break;										
									}	
									
								}

							}
							
						}				
					}		

				}
			
			}
			else
			{
				Log.e("KOOLJ_log", "There is no TEST to run");
			}

		}
	}
	
//Define Robotium keyword	
//===========================================================
	public void solo_assertCurrentActivity (String message, Class expectedClass, Solo solo) {
		
		//solo.assertCurrentActivity(message, expectedClass);
		Log.e("KOOLJ_assertCurrentActivity_", message);
	}
	public void solo_clickonbutton (String value, Solo solo){
		if(value.equals("0")) 
		{
			int value_1=Integer.parseInt(value);
			Log.e("KOOLJ_clickonbutton_", ""+value);
			solo.clickOnButton(value_1);
		} 
		else if(value.equals("1")) 
		{
			int value_1=Integer.parseInt(value);
			Log.e("KOOLJ_clickonbutton_", ""+value);
			solo.clickOnButton(value_1);
		} 
		else if(value.equals("2")) 
		{
			int value_1=Integer.parseInt(value);
			Log.e("KOOLJ_clickonbutton_", ""+value);
			solo.clickOnButton(value_1);
		}
		else
		{	
			Log.e("KOOLJ_clickonbutton_", ""+value);
			solo.clickOnButton(value);
			
		}
	}
	public void asl_screenshot (Float value1,  Float value2, Solo solo){
		
		//solo.clickOnScreen(value1,value2);
		//ArrayList[] Act2Report = solo.getAllOpenedActivities();
		//for (int i_act = 0; i_act < Act2Report.length; i_act++)
		//{
		//	Log.e("KoolJ_at", Act2Report[i_act].toString());
		//}

/*
		Log.e("KoolJ_screenshot", "SHOT!");
        String imageCapturedPath = ""; 
        try {
			if (RobotiumTest.aslProvider.isAvailable())
			{
				imageCapturedPath = RobotiumTest.aslProvider.takeScreenshot();
			}
			else
			{
				Log.e("KoolJ_ASL", "service isn't ready.");
			}
		} catch (RemoteException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		finally {
			Log.e("Kookj", "---------->" + imageCapturedPath);
		}
*/		
        
	}
	public void solo_enterkey (int text, String value, Solo solo){
		Log.e("KOOLJ_entertext_", value);
		solo.enterText(text, value);
	}
	public void solo_screenshot (Solo solo){
	
	}
	public void solo_back (Solo solo){
		Log.e("KOOLJ_goback2_", "goBack");
		solo.goBack();
	}
	public void solo_searchtext (String value, Solo solo){
		boolean value_expected = true;
		boolean value_actual = solo.searchText(value);
		Log.e("KOOLJ_SEARCHTEXT: "+value, ""+value_actual);
		KOOLJ_log=KOOLJ_log+"\n"+"SEARCH TEXT "+ "'" + value + "'" + " is "+value_actual;
	}
	public void solo_key (int value, Solo solo){
		Log.e("KOOLJ_sendKey_", ""+value);
		solo.sendKey(value);

	}	
	public void solo_sleep (int value, Solo solo){
		Log.e("KOOLJ_sleep_", ""+value);
		solo.sleep(value);
		
	}
	
//===========================================================
	//Load EXCEL file
	public Object[][] CreateDataFromCSV(String file_xls) { 
		//Checking environment SDCARD
		boolean mExternalStorageAvailable = false;
		boolean mExternalStorageWriteable = false;
		String state = Environment.getExternalStorageState();

		if (Environment.MEDIA_MOUNTED.equals(state)) {
		    // We can read and write the media
		    mExternalStorageAvailable = mExternalStorageWriteable = true;
		} else if (Environment.MEDIA_MOUNTED_READ_ONLY.equals(state)) {
		    // We can only read the media
		    mExternalStorageAvailable = true;
		    mExternalStorageWriteable = false;
		} else {
		    // Something else is wrong. It may be one of many other states, but all we need
		    //  to know is we can neither read nor write
		    mExternalStorageAvailable = mExternalStorageWriteable = false;
		}
		
		//PATH
		File rootsd = Environment.getExternalStorageDirectory();
	    File dcim = new File(rootsd.getAbsolutePath() + "/DCIM/DFRS");
		file_xls = dcim + file_xls;
		
		//Start to open to read file
		File DatatestExcel = new File(file_xls); 
		HSSFWorkbook workbook; 
		String[][] data = null; 
		FileInputStream stream = null;
		Log.e("XLS_load", file_xls);
		KOOLJ_log=KOOLJ_log+"\n"+"XLS_load" + file_xls;
		try { 
			stream = new FileInputStream(DatatestExcel); 
			workbook = new HSSFWorkbook(stream); 
			HSSFSheet sheet = workbook.getSheetAt(0); 
			int rows = sheet.getLastRowNum() + 1; 
			short cells = sheet.getRow(0).getLastCellNum(); 
			data = new String[rows][cells]; 
			List<String> list = new ArrayList<String>(); 
		
			for (int i = 0; i < rows; i++) { 
				HSSFRow row = sheet.getRow(i); 
					for (short j = 0; j < cells; j++) { 
						HSSFCell cell = row.getCell(j); 
						String value = null; 
						if (cell != null) { 
							value = cellToString(cell); 
						} 
						data [i][j] = value; 
					} 
			} 
		} 
		catch (FileNotFoundException e) { 
		// TODO Auto-generated catch block
			Log.e("XLS_notfound", e.fillInStackTrace().toString());
			e.printStackTrace(); 
		} 
		catch (IOException e) { 
		// TODO Auto-generated catch block 
			Log.e("Catch_IO_", e.fillInStackTrace().toString());
			e.printStackTrace(); 
		}
		finally {
			//close file
			if (stream != null) {
				try {
					stream.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		}
		return data; 
	}
	//Verify Excel results 
	public static String cellToString(HSSFCell cell) { 
		int type = cell.getCellType(); 
		Object result; 
		switch (type) { 
			case HSSFCell.CELL_TYPE_NUMERIC: // 0 
			result = cell.getNumericCellValue(); 
			break; 
			case HSSFCell.CELL_TYPE_STRING: // 1 
			result = cell.getStringCellValue(); 
			break; 
			case HSSFCell.CELL_TYPE_FORMULA: // 2 
			result = cell.getStringCellValue(); 
			//throw new RuntimeException("We can't evaluate formulas in Java"); 
			case HSSFCell.CELL_TYPE_BLANK: // 3 
			result = ""; 
			break; 
			case HSSFCell.CELL_TYPE_BOOLEAN: // 4 
			result = cell.getBooleanCellValue(); 
			break; 
			case HSSFCell.CELL_TYPE_ERROR: // 5 
			throw new RuntimeException("This cell has an error"); 
			default: 
			throw new RuntimeException("We don't support this cell type: " + type); 
		} 
		return result.toString(); 
	}

	//Download EXCEL file from HTTP into DCIM
	public void URLfile(String urlfile, String download_file)
	{
		Log.e("KOOLJ_downloading", download_file);
		KOOLJ_log=KOOLJ_log+"\n"+"KOOLJ_downloading" + download_file;
		try {
	        //set the download URL (not including file), a url that points to a file on the internet
	        URL url = new URL(urlfile);
			//Log.e("KOOLJ_FULL", urlfile+download_file);
	        //create the new connection
	        HttpURLConnection urlConnection = (HttpURLConnection) url.openConnection();

	        //set up some things on the connection
	        urlConnection.setRequestMethod("GET");
	        urlConnection.setDoOutput(true);
			
	        //and connect!
	        urlConnection.connect();

	        //set the path where we want to save the file
	        //in this case, going to save it on the root directory of the
	        //sd card.
	        File SDCardRoot = Environment.getExternalStorageDirectory();
	        
	        //create a new file, specifying the path, and the filename
	        //which we want to save the file as.
			//PATH
		    File dcim = new File(SDCardRoot.getAbsolutePath() + "/DCIM/DFRS");
	        File file = new File(dcim,download_file);

	        //this will be used to write the downloaded data into the file we created
	        FileOutputStream fileOutput = new FileOutputStream(file);

	        //this will be used in reading the data from the internet
	        InputStream inputStream = urlConnection.getInputStream();

	        //this is the total size of the file
	        int totalSize = urlConnection.getContentLength();
	        //Log.e("KOOLJ_"+download_file, ""+totalSize);
			
	        //variable to store total downloaded bytes
	        int downloadedSize = 0;

	        //create a buffer...
	        byte[] buffer = new byte[1024];
	        int bufferLength = 0; //used to store a temporary size of the buffer

	        //now, read through the input buffer and write the contents to the file
	        while ( (bufferLength = inputStream.read(buffer)) > 0 ) 
			{
	                //add the data in the buffer to the file in the file output stream (the file on the sd card
	                fileOutput.write(buffer, 0, bufferLength);
	                //add up the size so we know how much is downloaded
	                downloadedSize += bufferLength;
	                //this is where you would do something to report the prgress, like this maybe
	                updateProgress(downloadedSize, totalSize);
					//Log.e("KOOLJ_startdownload", download_file);
	        }
	        //close the output stream when done
			
			file_download_done++;
	        fileOutput.close();
			urlConnection.disconnect();

		//catch some possible errors...
		} catch (MalformedURLException e) {
		        e.printStackTrace();
		} catch (IOException e) {
		        e.printStackTrace();
		}
	}

	private void updateProgress(int downloadedSize, int totalSize) {
		String downprogress_var;
		//Log.e("KOOLJ_loading...", Long.toString((downloadedSize/totalSize)*100)+"%");
		KOOLJ_log=KOOLJ_log+"\n"+"Downloading status... "+Long.toString((downloadedSize/totalSize)*100)+"%"; 
	} 

}
