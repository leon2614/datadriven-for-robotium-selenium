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
import android.view.View;

import com.jayway.android.robotium.solo.Solo;

public class KoolJ_datadriven {
	
	String test_xls;
	String suite_xls;
	String map_xls;
	String batch_xls;
	String config_xls;
	String KOOLJ_log;
	String project_folder = "";
	String[][] outputReport;
	Object[][] data_suite;
	Object[][] data_test;
	Object[][] data_key;
	Object[][] data_url_batch;
	int read_first = 1;
	int file_download_done = 0;
	
	long starttime = 0;
	long endtime = 0;
	long elapsedtime = 0;	
	
	
	//Open CONFIG to BATCH,SUITE,TEST files
	public void openconfig(String config_xls, Solo solo){
		Object[][] data_batch = CreateDataFromCSV(config_xls);

		//check NULL data_batch
		if (data_batch == null) 
		{
			Log.e("KOOLJ_log", "DATA IS NOT AVAIL");
			//KOOLJ_log=KOOLJ_log+"\n"+"DATA IS NULL";
		}
		else
		{
			Log.e("KOOLJ_log", "DATA IS AVAIL");

			//Get project folder
			if (data_batch[1][0].toString().equals("project_folder"))
			{
				project_folder = data_batch[1][1].toString();
			}
			else 
			{
				Log.e("KOOLJ_log", "THERE IS NO PROJECT FOLDER");
			}	
			
			//if files from HTTP, download them
			if (!project_folder.equals(""))
			{
				file_download_done = 1;
				data_url_batch = CreateDataFromCSV("/url_batch.xls");
				for (int i_d=0; i_d< data_url_batch.length; i_d++)
				{					
					if (data_url_batch[i_d][1].toString().equals("yes"))
					{
						URLfile(data_url_batch[i_d][2].toString(),data_url_batch[i_d][0].toString());
					}
					else 
					{
						file_download_done = 3;
					}	
					
				}
			}
			else
			{	
				Log.e("KOOLJ_log", "THERE IS NO PROJECT FOLDER");
			}	
			//Find to run BATCH	
			if (!project_folder.equals(""))
			{
				if (file_download_done > 1)
				{
					//Find to run SUITE
					String data_suite_var="/batch.xls";
					//KOOLJ_log=KOOLJ_log+"\n"+"RUN BATCH: "+ data_suite_var;
					Log.e("KOOLJ_BATCH: ", data_suite_var);
					data_suite = CreateDataFromCSV(data_suite_var);
				}
			}
			else
			{	
				Log.e("KOOLJ_log", "THERE IS NO PROJECT FOLDER");
			}
			
			//Find to run TEST
			if (file_download_done > 2)
			{
				for (int ii=0; ii< data_suite.length; ii++)
				{
					
					String data_test_var="/" + data_suite[ii][0].toString() + ".xls";
					//KOOLJ_log=KOOLJ_log+"\n"+"RUN SUITE:______ "+ data_test_var;
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
						String[] key_if=new String[data_key.length];
						String[] key_endif=new String[data_key.length];
						String[] key_else=new String[data_key.length];

						int for_count = 0;
						int for_count_backward = 0;
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
						int key_stepstart = 0;
						int key_stepend = 0;
						int key_stepacc = 0;
						int if_step = 0;
						int if_step_backward = 0;
						int else_step_backward = 0;
						int if_count = 0;
						int if_count_backward = 0;
						int if_logic = 0;
						int if_located = 0;
						
						String key_ifstart = "";
						String key_ifend = "";
												
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
								key_endfor[for_step_backward] = ""+iiii;
							}
						}
						
						//Count STORE
						for (int iiii=iiii_label; iiii< data_key.length; iiii++)
						{
							String key_target = data_key[iiii][1].toString();
							if (key_target.equals("store"))
							{
								varstore_count++;
							}
						}
						
						//Store values of STORE if have
						Object[][] varstore_kv=new Object[varstore_count][2];
						for (int iiii=iiii_label; iiii< data_key.length; iiii++)
						{
							String key_target = data_key[iiii][1].toString();
							if (key_target.equals("store"))
							{
								varstore_kv[varstore_step][0] = data_key[iiii][2].toString();
								varstore_kv[varstore_step][1] = data_key[iiii][3].toString();
								varstore_step++;									
							}
						}
																				
						//Count IF..ELSE..ENDIF if have
						for (int iiii=iiii_label; iiii< data_key.length; iiii++)
						{
							String key_target = data_key[iiii][1].toString();
							if (key_target.equals("if"))
							{
								if_located = iiii;
								key_if[if_step] = ""+if_located;
								//Log.e("KOOLJ_if_", key_if[if_step]);
								if_step++;
								if_step_backward = if_step;
								else_step_backward = if_step;
								
							}
							else if (key_target.equals("endif"))
							{
								if_located = iiii;
								if_step_backward--;
								key_endif[if_step_backward] = ""+if_located;
								//Log.e("KOOLJ_end_", key_endif[if_step_backward]);
							}	
							else if (key_target.equals("else"))
							{
								if_located = iiii;
								else_step_backward--;
								key_else[else_step_backward] = ""+if_located;
								//Log.e("KOOLJ_else_", key_else[else_step_backward] );
							}	
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
							else if(key_target.equals("if"))
							{
								int var_if = 0;
								int i_stepstart = 0;
								String key_logic = "";
								String var_temp = "";
								for (int i = 0; i< varstore_kv.length; i++)
								{
									var_temp = varstore_kv[i][0].toString();
									key_logic = data_key[iiii][3].toString();
									if (var_temp.equals(data_key[iiii][2].toString()))
									{									
										key_ifstart = varstore_kv[i][1].toString();
										i_stepstart = i;
										var_if++;
									}
									else if (var_temp.equals(data_key[iiii][4].toString()))
									{									
										key_ifend = varstore_kv[i][1].toString();
										var_if++;
									}
									else
									{
										if ( var_if > 1)
										{
											break;
										}
									}
								}
								
								//Change step if logic on IF..ENDIF valid
								if ( var_if == 0)
								{
									String key_wait = data_key[iiii][2].toString();
									String key_waitval1 = "0";
									String key_waitval2 = "0";
									String key_waitval3 = "0";
									if (key_wait.equals("waitForActivity"))
									{
										int key_waitval2_var = Integer.parseInt(key_waitval2);
										key_waitval1 = data_key[iiii][3].toString();
										key_waitval2 = data_key[iiii][4].toString();
										if (solo_waitForActivity(key_waitval1, key_waitval2_var, solo))
										{
											if_logic = 0;
										}
										else
										{
											if_logic = 1;
											if (!key_else[if_count].equals(null))
											{
												iiii_label = Integer.parseInt(key_else[if_count]);	
												iiii = iiii_label;
											}
											else
											{
												iiii_label = Integer.parseInt(key_endif[if_count]);	
												iiii = iiii_label;												
											}
										}
									}
									else if (key_wait.equals("waitForView"))
									{
										key_waitval2 = data_key[iiii][4].toString();
										int key_waitval2_var = Integer.parseInt(key_waitval2);
										key_waitval1 = data_key[iiii][3].toString();
										if (solo_waitForView(key_waitval1, key_waitval2_var, solo))
										{
											if_logic = 0;
										}
										else
										{
											if_logic = 1;
											if (!key_else[if_count].equals(null))
											{
												iiii_label = Integer.parseInt(key_else[if_count]);	
												iiii = iiii_label;
											}
											else
											{
												iiii_label = Integer.parseInt(key_endif[if_count]);	
												iiii = iiii_label;												
											}
										}
									}
									else if (key_wait.equals("waitForText"))
									{
										key_waitval1 = data_key[iiii][3].toString();
										key_waitval2 = data_key[iiii][4].toString();
										key_waitval3 = data_key[iiii][5].toString();
										int key_waitval2_var = Integer.parseInt(key_waitval2);
										long key_waitval3_var = Long.valueOf(key_waitval3);
										if (solo_waitForText(key_waitval1, key_waitval2_var, key_waitval3_var , solo))
										{
											if_logic = 0;
										}
										else
										{
											if_logic = 1;
											if (!key_else[if_count].equals(null))
											{
												iiii_label = Integer.parseInt(key_else[if_count]);	
												iiii = iiii_label;
											}
											else
											{
												iiii_label = Integer.parseInt(key_endif[if_count]);	
												iiii = iiii_label;												
											}
										}
									}
									else if (key_wait.equals("waitForDialogToClose"))
									{
										long key_waitval1_var = Long.valueOf(key_waitval1);
										if (solo_waitForDialogToClose (key_waitval1_var, solo))
										{
											if_logic = 0;
										}
										else
										{
											if_logic = 1;
											if (!key_else[if_count].equals(null))
											{
												iiii_label = Integer.parseInt(key_else[if_count]);	
												iiii = iiii_label;
											}
											else
											{
												iiii_label = Integer.parseInt(key_endif[if_count]);	
												iiii = iiii_label;												
											}
										}
									}
									else if (key_wait.equals("searchText"))
									{
										if (solo_searchtext(key_waitval1, solo))
										{
											if_logic = 0;
										}
										else
										{
											if_logic = 1;
											if (!key_else[if_count].equals(null))
											{
												iiii_label = Integer.parseInt(key_else[if_count]);	
												iiii = iiii_label;
											}
											else
											{
												iiii_label = Integer.parseInt(key_endif[if_count]);	
												iiii = iiii_label;												
											}
										}
									}
									else
									{
										//go to ENDIF
										if_logic = 1;
										iiii_label = Integer.parseInt(key_endif[if_count]);	
										iiii = iiii_label;
									}	
								}
								else
								{
									//validate each logic parameters
									int var_key_ifstart = 0;
									int var_key_ifend = 0;
									
									if ( key_logic.equals("''="))
									{	
										if (!key_ifstart.equals(key_ifend))
										{
											if_logic = 1;
											if (!key_else[if_count].equals(null))
											{
												iiii_label = Integer.parseInt(key_else[if_count]);	
												iiii = iiii_label;
											}
											else
											{
												iiii_label = Integer.parseInt(key_endif[if_count]);	
												iiii = iiii_label;												
											}	
										}
										else
										{
											if_logic = 0;												
										}											
									}
									else if ( key_logic.equals("#"))
									{
										if (key_ifstart.equals(key_ifend))
										{
											if_logic = 1;
											if (!key_else[if_count].equals(null))
											{
												iiii_label = Integer.parseInt(key_else[if_count]);	
												iiii = iiii_label;
											}
											else
											{
												iiii_label = Integer.parseInt(key_endif[if_count]);	
												iiii = iiii_label;	
											}
										}
										else
										{
											if_logic = 0;
										}	
									}
									else if ( key_logic.equals(">"))
									{
										var_key_ifstart = Integer.parseInt(key_ifstart);
										var_key_ifend = Integer.parseInt(key_ifend);
										if (var_key_ifstart < var_key_ifend)
										{
											if_logic = 1;
											if (!key_else[if_count].equals(null))
											{
												iiii_label = Integer.parseInt(key_else[if_count]);	
												iiii = iiii_label;
											}
											else
											{
												iiii_label = Integer.parseInt(key_endif[if_count]);	
												iiii = iiii_label;												
											}	
										}
										else
										{
											if_logic = 0;
										}
									}
									else if ( key_logic.equals("<"))
									{
										var_key_ifstart = Integer.parseInt(key_ifstart);
										var_key_ifend = Integer.parseInt(key_ifend);
										if (var_key_ifstart > var_key_ifend)
										{
											if_logic = 1;
											if (!key_else[if_count].equals(null))
											{
												iiii_label = Integer.parseInt(key_else[if_count]);	
												iiii = iiii_label;
											}
											else
											{
												iiii_label = Integer.parseInt(key_endif[if_count]);	
												iiii = iiii_label;												
											}	
										}
										else
										{
											if_logic = 0;
										}
									}
									else if ( key_logic.equals("<="))
									{
										var_key_ifstart = Integer.parseInt(key_ifstart);
										var_key_ifend = Integer.parseInt(key_ifend);
										if (var_key_ifstart > var_key_ifend)
										{
											if_logic = 1;
											if (!key_else[if_count].equals(null))
											{
												iiii_label = Integer.parseInt(key_else[if_count]);	
												iiii = iiii_label;
											}
											else
											{
												iiii_label = Integer.parseInt(key_endif[if_count]);	
												iiii = iiii_label;												
											}	
										}
										else
										{
											if_logic = 0;
										}
									}
									else if ( key_logic.equals(">="))
									{
										var_key_ifstart = Integer.parseInt(key_ifstart);
										var_key_ifend = Integer.parseInt(key_ifend);
										if (var_key_ifstart < var_key_ifend)
										{
											if_logic = 1;
											if (!key_else[if_count].equals(null))
											{
												iiii_label = Integer.parseInt(key_else[if_count]);	
												iiii = iiii_label;
											}
											else
											{
												iiii_label = Integer.parseInt(key_endif[if_count]);	
												iiii = iiii_label;												
											}	
										}
										else
										{
											if_logic = 0;
										}
									}
									else
									{
										iiii_label = Integer.parseInt(key_endif[if_count]);	
										iiii = iiii_label;
									}
								}
								if_count++;	
								if_count_backward = if_count;
							}
							else if(key_target.equals("else"))
							{	
								if_count_backward--;
								if (if_logic == 0)
								{
									iiii_label = Integer.parseInt(key_endif[if_count_backward]);	
									iiii = iiii_label;
								}
							}
							else if(key_target.equals("for"))
							{
								int var_for = 0;
								int i_stepstart = 0;
								//Get values and compare logic
								for (int i = 0; i< varstore_kv.length; i++)
								{
									String var_temp = varstore_kv[i][0].toString();
									if (var_temp.equals(data_key[iiii][2].toString()))
									{									
										key_stepstart = Integer.parseInt(varstore_kv[i][1].toString());
										i_stepstart = i;
										var_for++;
									}
									else if (var_temp.equals(data_key[iiii][4].toString()))
									{									
										key_stepend = Integer.parseInt(varstore_kv[i][1].toString());
										var_for++;
									}
									else if (var_temp.equals(data_key[iiii][3].toString()))
									{									
										key_stepacc = Integer.parseInt(varstore_kv[i][1].toString());
										var_for++;
									}
									else
									{
										if ( var_for > 2)
										{
											break;
										}
									}
								}
								
								//Change step if logic on FOR..ENDFOR valid
								if ( var_for == 0)
								{
									//go to ENDFOR
									iiii_label = Integer.parseInt(key_endfor[for_count]);	
									iiii = iiii_label;
								}
								else
								{
									if ( key_stepstart <= key_stepend)
									{	
										//go to FOR
										iiii_label = Integer.parseInt(key_for[for_count]);	
										iiii = iiii_label;	
										key_stepstart = key_stepstart + key_stepacc;
										varstore_kv[i_stepstart][1] = ""+key_stepstart;
									}
									else
									{
										//go to ENDFOR
										iiii_label = Integer.parseInt(key_endfor[for_count]);	
										iiii = iiii_label;
									}	
								}
								for_count++;	
								for_count_backward = for_count;
							}
							else if(key_target.equals("endfor"))
							{
								for_count_backward--;
								iiii_label = Integer.parseInt(key_for[for_count_backward].toString());	
								iiii = iiii_label - 1;	
							}
							else if(key_target.equals("store"))
							{
								//Search to change the VAR
								for (int iz = 0; iz< varstore_kv.length; iz++)
								{
									String var_temp = varstore_kv[iz][0].toString();
									if (var_temp.equals(data_key[iiii][2].toString()))
									{
										String var_temp3 = " ";
										for (int ix = 0; ix< varstore_kv.length; ix++)
										{
											String var_temp2 = varstore_kv[ix][0].toString();
											if (var_temp2.equals(data_key[iiii][3].toString()))
											{
												var_temp3 = varstore_kv[ix][1].toString();
												varstore_kv[iz][1] = varstore_kv[ix][1].toString();
												break;
											}
											else
											{
												varstore_kv[iz][1] = data_key[iiii][3].toString();
											}
										}
									}
								}
							}	
							else if(key_target.equals("echo"))
							{
								int echo_in = 0;
								for (int i = 0; i< varstore_kv.length; i++)
								{
									String var_temp = varstore_kv[i][0].toString();
									if (var_temp.equals(data_key[iiii][2].toString()))
									{									
										Log.e("KOOLJ_ECHO_" + varstore_kv[i][0].toString(), varstore_kv[i][1].toString());
										echo_in = 1;
										break;
									}
								}
								if (echo_in == 0)
									Log.e("KOOLJ_ECHO_", data_key[iiii][2].toString());
							}							
							else if(key_target.equals("waitForActivity"))
							{
								String key_value1 = data_key[iiii][2].toString();
								int key_value2 = Integer.parseInt(data_key[iiii][3].toString());
								solo_waitForActivity (key_value1, key_value2, solo);
							}
							else if(key_target.equals("screenshot"))
							{
								String key_value = data_key[iiii][2].toString();
								solo_screenshot(solo, key_value);
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
							else if(key_target.equals("waitForDialogToClose"))
							{
								long key_value = Long.valueOf(data_key[iiii][3].toString());
								solo_waitForDialogToClose (key_value, solo);
								
							}
							else if(key_target.equals("waitForText"))
							{
								String key_text = data_key[iiii][2].toString();
								int key_minimumNumberOfMatches = Integer.parseInt(data_key[iiii][3].toString());
								long key_timeout = Long.valueOf(data_key[iiii][4].toString());
								solo_waitForText (key_text, key_minimumNumberOfMatches, key_timeout , solo);
								
							}
							else if(key_target.equals("waitForView"))
							{
								String key_view = data_key[iiii][2].toString();
								int key_timeout = Integer.parseInt(data_key[iiii][3].toString());
								solo_waitForView (key_view, key_timeout, solo);
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
	
//Define Robotium keywords
//===========================================================
	public boolean solo_waitForView (String view, int timeout, Solo solo) {
		Object[] view_arr = solo.getCurrentViews().toArray();
		boolean waitForView_status =  false;
		for (int i = 0; i < view_arr.length; i++ )
		{
			String current_view = view_arr[i].toString();
			if (current_view.equals(view))
			{
				Log.e("KOOLJ_waitForView_"+solo.getCurrentViews().get(i), ""+timeout);
				starttime = System.currentTimeMillis();
				waitForView_status =  solo.waitForView(solo.getViews().get(i), timeout, true);
				endtime = System.currentTimeMillis();
				elapsedtime = (endtime - starttime)/1000;
				break;
			}
		}
		return waitForView_status;
	}
	public boolean solo_waitForText (String text, int minimumNumberOfMatches,long timeout, Solo solo) {
		starttime = System.currentTimeMillis();
		boolean waitForText_status =  solo.waitForText(text, minimumNumberOfMatches, timeout, true);
		endtime = System.currentTimeMillis();
		elapsedtime = (endtime - starttime)/1000;
		Log.e("KOOLJ_waitForText_"+text, ""+timeout);
		return waitForText_status;
	}
	public boolean solo_waitForDialogToClose (long timeout, Solo solo) {
		starttime = System.currentTimeMillis();
		boolean waitForDialogToClose_status =  solo.waitForDialogToClose(timeout);
		endtime = System.currentTimeMillis();
		elapsedtime = (endtime - starttime)/1000;
		Log.e("KOOLJ_waitForDialogToClose_", ""+timeout);
		return waitForDialogToClose_status;
	}
	public boolean solo_waitForActivity (String name, int timeout, Solo solo) {
		Object[] activity_arr = solo.getAllOpenedActivities().toArray();
		boolean waitForActivity_status =  false;
		for (int i = 0; i < activity_arr.length; i++ )
		{
			String current_activity = activity_arr[i].toString();
			if (current_activity.equals(name))
			{
				Log.e("KOOLJ_waitForActivity_"+solo.getAllOpenedActivities().get(i), ""+timeout);
				starttime = System.currentTimeMillis();
				waitForActivity_status =  solo.waitForActivity(name, timeout);
				endtime = System.currentTimeMillis();
				elapsedtime = (endtime - starttime)/1000;
				break;
			}
		}
		return waitForActivity_status;
	}
	public void solo_clickonbutton (String value, Solo solo)
	{
		if(value.equals("0")) 
		{
			int value_1=Integer.parseInt(value);
			Log.e("KOOLJ_clickonbutton_", ""+value);
			starttime = System.currentTimeMillis();
			solo.clickOnButton(value_1);
			endtime = System.currentTimeMillis();
			elapsedtime = (endtime - starttime)/1000;
		} 
		else if(value.equals("1")) 
		{
			int value_1=Integer.parseInt(value);
			Log.e("KOOLJ_clickonbutton_", ""+value);
			starttime = System.currentTimeMillis();
			solo.clickOnButton(value_1);
			endtime = System.currentTimeMillis();
			elapsedtime = (endtime - starttime)/1000;
		} 
		else if(value.equals("2")) 
		{
			int value_1=Integer.parseInt(value);
			Log.e("KOOLJ_clickonbutton_", ""+value);
			starttime = System.currentTimeMillis();
			solo.clickOnButton(value_1);
			endtime = System.currentTimeMillis();
			elapsedtime = (endtime - starttime)/1000;
		}
		else
		{	
			Log.e("KOOLJ_clickonbutton_", ""+value);
			starttime = System.currentTimeMillis();
			solo.clickOnButton(value);
			endtime = System.currentTimeMillis();
			elapsedtime = (endtime - starttime)/1000;
		}
	}
    public void solo_enterkey (int text, String value, Solo solo){
		Log.e("KOOLJ_entertext_", value);
		starttime = System.currentTimeMillis();
		solo.enterText(text, value);
		endtime = System.currentTimeMillis();
		elapsedtime = (endtime - starttime)/1000;
	}
	public void solo_screenshot (Solo solo, String name){
		starttime = System.currentTimeMillis();
		Screenshot takeSS = new Screenshot();
		try 
		{
			Log.e("KoolJ_getScreenshot_"+solo.getViews().get(0), name);

			takeSS.takeScreenShot(solo.getViews().get(0), name);
			endtime = System.currentTimeMillis();
			elapsedtime = (endtime - starttime)/1000;
		}
		catch (Exception e) 
		{
			Log.e("KoolJ_errorScreenshot", e.getMessage());
		}
	}
	public void solo_back (Solo solo){
		Log.e("KOOLJ_goback2_", "goBack");
		starttime = System.currentTimeMillis();
		solo.goBack();
		endtime = System.currentTimeMillis();
		elapsedtime = (endtime - starttime)/1000;
	}
	public boolean solo_searchtext (String value, Solo solo){
		boolean value_expected = true;
		starttime = System.currentTimeMillis();
		boolean value_actual = solo.searchText(value, 0, true, true);
		endtime = System.currentTimeMillis();
		elapsedtime = (endtime - starttime)/1000;
		Log.e("KOOLJ_SEARCHTEXT: "+value, ""+value_actual);
		//KOOLJ_log=KOOLJ_log+"\n"+"SEARCH TEXT "+ "'" + value + "'" + " is "+value_actual;
		return value_actual;
	}
	public void solo_key (int value, Solo solo){
		Log.e("KOOLJ_sendKey_", ""+value);
		starttime = System.currentTimeMillis();
		solo.sendKey(value);
		endtime = System.currentTimeMillis();
		elapsedtime = (endtime - starttime)/1000;
	}	
	public void solo_sleep (int value, Solo solo){
		Log.e("KOOLJ_sleep_", ""+value);
		starttime = System.currentTimeMillis();
		solo.sleep(value);
		endtime = System.currentTimeMillis();
		elapsedtime = (endtime - starttime)/1000;
		//Log.e("KOOLJ_elapsedtime_", ""+elapsedtime);
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
	    if (file_xls.equals("/config.xls"))
		{
			dcim = new File(rootsd.getAbsolutePath() + "/DCIM/DFRS");
		}
		else
		{
			dcim = new File(rootsd.getAbsolutePath() + "/DCIM/DFRS/"+project_folder);
		}
		file_xls = dcim + file_xls;
		
		//Start to open to read file
		File DatatestExcel = new File(file_xls); 
		HSSFWorkbook workbook; 
		String[][] data = null; 
		FileInputStream stream = null;
		Log.e("XLS_load", file_xls);
		//KOOLJ_log=KOOLJ_log+"\n"+"XLS_load" + file_xls;
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
		    File dcim = new File(SDCardRoot.getAbsolutePath() + "/DCIM/DFRS"+"/"+project_folder);
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
		
	//Building @Dataprovider named "DataTestMSSQL" from Microsoft SQL Server 
	//@BeforeTest
	//@DataProvider(name = "DataTestMSSQL") 
/*
	public String [][] SQL_Data() 
	{ 
	int rowCount = 0; 
	int columnCount = 0; 
	String myData [][] = null; 

	try 
	{ 
		//System.out.println("Successfully connected to KoolJ_MSSQL...");
		Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver"); 
		String url = "databaseName=internet_vn;integratedSecurity=false;selectMethod=direct"; 
	
		Connection con = DriverManager.getConnection("jdbc:sqlserver://svfpt04;"+url,"sa","sa"); 

		// Execute the SQL statement 
		//Statement stmt = con.createStatement(ResultSet.TYPE_SCROLL_INSENSITIVE,ResultSet.CONCUR_UPDATABLE); 
		Statement stmt = con.createStatement(ResultSet.TYPE_SCROLL_INSENSITIVE,ResultSet.CONCUR_READ_ONLY);
		ResultSet resultSet = stmt.executeQuery("Select top 3 substring(ITEM_NO,1,4),substring(ITEM_NO,5,5) from Cop_sku");//("EXEC TestReplayData2 '20080109 08:00:00', '20080109 16:30:00'"); 
		
		
		CallableStatement Proc_State=con.prepareCall("{ call TestReplayData2(?,?)}");
		Proc_State.setString(1, "20080109 08:00:00");
		Proc_State.setString(2, "20080109 16:30:00");
		
		// Get Column count 
		ResultSetMetaData resultSet_metaData= resultSet.getMetaData(); 
		columnCount = resultSet_metaData.getColumnCount(); 

		// Get Row Count 
		while( resultSet.next() ) 
		rowCount++; 

		//Initialize data structure 
		myData = new String [rowCount][columnCount]; 

		resultSet.beforeFirst(); 


		//populate data structure 
		for(int row=0; row<rowCount; row++) 
		{ 
			resultSet.next(); 
			for(int col=1; col <=columnCount; col++) 
			myData[row][col-1] = resultSet.getString(col); 
			//System.out.println("");
		} 
		resultSet.close();
        stmt.close();
        con.close();
	}
	catch (Exception e) 
	{ 
		e.printStackTrace(); 
	} 
	return myData; 
	}
*/
/*
	//Building @Dataprovider named "DataTestMySQL" from MySQL 5.4 
	//@BeforeTest 
	//@DataProvider(name = "DataTestMySQL") 
	public String [][] MySQL_Data() 
	{
	    Connection con = null;
	    int rowCount = 0; 
		int columnCount = 0;
		String myData [][] = null; 
	    try {
	      Class.forName("com.mysql.jdbc.Driver").newInstance();
	      con = DriverManager.getConnection("jdbc:mysql://localhost:3306/fso_timesheet","root","12345");

	      if(!con.isClosed())
	        System.out.println("Successfully connected to KoolJ_MySQL...");
	      	Statement stmt = con.createStatement(ResultSet.TYPE_SCROLL_INSENSITIVE,ResultSet.CONCUR_READ_ONLY);
	      	ResultSet resultSet = stmt.executeQuery("select * from add_new");
	      	
	      	ResultSetMetaData resultSet_metaData= resultSet.getMetaData(); 
			columnCount = resultSet_metaData.getColumnCount(); 

			// Get Row Count 
			while( resultSet.next() ) 
			rowCount++; 

			//Initialize data structure 
			myData = new String [rowCount][columnCount]; 
			resultSet.beforeFirst(); 
			
			int col;
			//populate data structure 
			for(int row=0; row<rowCount; row++) 
			{ 
				resultSet.next(); 
				for(col=1; col <=columnCount; col++) 
				{
					myData[row][col-1] = resultSet.getString(col); 
					//System.out.println("DATA-------recordnumber--"+row+": "+myData[row][col-1]);
				}
			} 
			
			resultSet.close();
	        stmt.close();
	        con.close();
	        
	    } catch(Exception e) {
	      System.err.println("Exception: " + e.getMessage());
	    } finally {
	      try {
	        if(con != null)
	          con.close();
	      } catch(SQLException e) {}
	    }
	  	return myData; 
	}
*/	
}
