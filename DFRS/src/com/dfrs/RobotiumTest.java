package com.dfrs;

import android.test.ActivityInstrumentationTestCase2;
import android.test.suitebuilder.annotation.Smoke;
import android.util.Log;

import com.dfrs.R;
import com.jayway.android.robotium.solo.Solo;

@SuppressWarnings("unchecked")
public class RobotiumTest extends ActivityInstrumentationTestCase2
{
	private static final String	TARGET_PACKAGE_ID = "com.doximity.doximitydroid";
	private static final String LAUNCHER_ACTIVITY_FULL_CLASSNAME = "com.doximity.doximitydroid.app.screens.GenSplash";
	private	static Class<?>	launcherActivityClass;
	static
	{
		try
		{
			launcherActivityClass =	Class.forName(LAUNCHER_ACTIVITY_FULL_CLASSNAME);
		}
		catch (ClassNotFoundException e)
		{
			throw new RuntimeException(e);
		}
	}
	
	@SuppressWarnings("unchecked")
	public RobotiumTest() throws	ClassNotFoundException
	{
		super(TARGET_PACKAGE_ID,launcherActivityClass);
		//super("com.doximity.doximitydroid.app.screens", GenSplash.class);
	}
	public Solo solo;

	@Override 
	protected void setUp() throws Exception 
	{
		solo=new Solo(getInstrumentation(),	getActivity());
	}
	
	//******************************************************************************************
	//START DEVELOPING TESTCASE
	//******************************************************************************************
/*	
	@Smoke
	public void test1login()
	{
		//Click YES to login
		solo.clickOnButton("Yes"); 
		solo.enterText(0, "n@doc.com"); // hoan@doximity.com/aaaaaaaa, zz@dox.com/aaaaaaaa
		solo.enterText(1, "aaaaaaaa");
		solo.clickOnButton(0);
		//Confirm you are login
		solo.sleep(20000);
		boolean expected1 = true;
		boolean actual1 = solo.searchText("Lisa");
		assertEquals("You are now LOGGED IN", expected1, actual1); 
		Log.e("DONE LOGIN", "----------->Searching... "+actual1);
	}
*/
	@Smoke
	public void test0run()
	{
		
		//Call config CSV file
		KoolJ_datadriven Kjdriven=new KoolJ_datadriven();
		Kjdriven.openconfig("/config.xls","/output.xls", solo);
	}
/*	
	@Smoke
	public void test7viewfacility()
	{
		//Go to facility
		solo.sleep(5000);
		solo.sendKey(20);
		solo.sendKey(20);
		solo.sendKey(22);
		solo.sendKey(20);
		solo.sendKey(66);
		
		//Go to list
		solo.sleep(2000);
		solo.clickLongOnScreen(5,475);
		solo.sleep(20000);

	}
	*/
/*
	@Smoke
	public void test3findspecialty()
	{
		//back to main screen
		//solo.goBack();
		
		//Go to find
		solo.sleep(5000);
		solo.sendKey(20);
		solo.sendKey(66);
		
		//Go to my specialty
		solo.sendKey(66);
		
		//Verify
		solo.sleep(20000);
		boolean expected1 = true;
		boolean actual1 = solo.searchText("Ahmad");
		///assertEquals("The person on same specialty is: "+actual1, expected1, actual1); 
		Log.e("DONE FINDSPEC","----------->Searching... "+actual1);
	}
	*/
	/*
	public void testAboutBtn() {
		//Go to facility
		solo.sleep(5000);
		solo.sendKey(20);
		solo.sendKey(20);
		solo.sendKey(22);
		solo.sendKey(66);
		
		
		Log.i("LOGGER", ">>>>> Current Activity::::" + solo.getCurrentActivity().toString());
		
		// on click
		
		ImageView btnList = (ImageView) solo.getCurrentActivity().findViewById(R.id.img_List);
		
		
//		solo.getCurrentActivity().
		
		
		List<ImageView> imageViewList = solo.getCurrentImageViews();
		
		Log.e("LOGGER", ">>>>>>>>>> id of image view:::" + btnList.getId());
		
		int index = 0;
		for (ImageView item : imageViewList)
		{
			if (item.getId() == btnList.getId())
			{
				Log.e("LOGGER", "Equal=========" + index);
//				solo.clickOnImage(index);
				break;
			}
			
			Log.e("LOGGER", ">>> index:::" + index);
			
			index++;	
		}
		
		solo.clickOnImage(index);
		assertEquals(AcFuncFacilitiesMap.class, solo.getCurrentActivity().getClass());
		
//		solo.clickOnImage(btnList.getId());
		
		
	      // get a list of all ImageButtons on the current activity
	      
		
/*		  List btnList = solo.getCurrentImageButtons();
	      
	      for (int i = 0; i < btnList.size(); i++) {
	          
	    	  Log.e("Current BTN: ",solo.getCurrentImageButtons().get(i).toString());
	    	  Object btn = btnList.get(i);
	            // find button by id
	            if (btn.getId() == R.id.about_button) {
	                  // click on the button using index (not id !!!)
	                  solo.clickOnImageButton(i);
	                  // check if new activity is the 'About'
	                  assertEquals(About.class, solo.getCurrentActivity().getClass());
	                  Log.e("Current BTN: ",btn_var1);
	            } else {
	            	  Log.e("Current BTN: ",btn_var1);
	            }
	      }
	}
	*/

	/*
	@Smoke
	public void test6viewmap()
	{
		//Go to colleague
		solo.sleep(5000);
		solo.sendKey(20);
		solo.sendKey(22);
		solo.sendKey(66);
		
		//Go to map
		solo.sleep(2000);
		solo.clickOnScreen(280,50);
		solo.clickOnScreen(280,50);
		
		//Verify
		solo.sleep(20000);
		String btn_var1=solo.getCurrentImageButtons().get(0).toString();
		String btn_var2=solo.getCurrentImageButtons().get(1).toString();
		String btn_var3=solo.getCurrentImageButtons().get(2).toString();
		String btn_var4=solo.getCurrentImageButtons().get(3).toString();
		Log.e("Current BTN: ",btn_var1);
		Log.e("Current BTN: ",btn_var2);
		Log.e("Current BTN: ",btn_var3);
		Log.e("Current BTN: ",btn_var4);

		boolean expected1 = true;
		boolean actual1 = solo.searchText("Aaron");
		assertEquals("The person on map is: "+actual1, expected1, actual1); 
		Log.e("DONE FINDLOC","----------->Searching... "+actual1);
	}
	*/
	/*
	@Smoke
	public void test5doctext()
	{
		//Go to doctext
		solo.sleep(5000);
		solo.sendKey(20);
		solo.sendKey(22);
		solo.sendKey(22);
		solo.sendKey(66);
		
		//Accept to write
		solo.sleep(5000);
		solo.clickOnScreen(10, 310);
		//solo.clickOnMenuItem("newMsg");

		//For Shari
		solo.clickLongOnText("Shari Buckphone");
		
		//Go to text field
		solo.sendKey(20);
		solo.sendKey(66);
		
		//Typing
		solo.sleep(2000);
		solo.sendKey(20);
		solo.sendKey(66);
		solo.sendKey(20);
		solo.enterText(1, "Thank you for accepting me!"); 
		solo.sendKey(22);
		solo.sendKey(66);
				
		//Verify
		solo.sleep(20000);
		boolean expected1 = true;
		boolean actual1 = solo.searchText("Thank you for accepting me!");
		assertEquals("Doctext message is : "+actual1, expected1, actual1); 
		Log.e("DONE SENDDOC","----------->Searching... "+actual1);
	}
//	@Smoke
//	public void test4findlocation()
//	{
//		//back to main screen
//		//solo.goBack();
//		//solo.goBack();
//		//solo.goBack();
//		
//		//Go to find
//		solo.sleep(5000);
//		solo.sendKey(20);
//		solo.sendKey(66);
//		
//		//Go to advance search
//		solo.sendKey(20);
//		solo.sendKey(20);
//		solo.sendKey(20);
//		solo.sendKey(20);
//		solo.sendKey(20);
//		solo.sendKey(20);
//		solo.sendKey(66);
//		
//		//Go to location
//		solo.sleep(2000);
//		solo.sendKey(20);
//		solo.sendKey(66);
//		
//		//Go to nearby
//		solo.sleep(2000);
//		solo.sendKey(20);
//		solo.sendKey(66);
//		
//		//Add 30 radius
//		solo.sendKey(20);
//		solo.sendKey(20);
//		solo.sendKey(67);
//		solo.sendKey(67);
//		solo.sendKey(67);
//		//solo.enterText(2,"30");
//		
//		EditText edtText = (EditText) solo.getCurrentActivity().findViewById(R.id.edt_Nearby);
//		solo.enterText(edtText, "30");
//		
//		
//		//Go search		
//		solo.clickOnButton(3);
//		solo.sleep(2000);
//		solo.clickOnButton(0);
//		
//		//Verify
//		solo.sleep(20000);
//		boolean expected1 = true;
//		boolean actual1 = solo.searchText("Aaron");
//		assertEquals("The person on 30 radius far is: "+actual1, expected1, actual1); 
//		Log.e("DONE FINDLOC","----------->Searching... "+actual1);
//	}
	*/


/*	public void test10logout()
	{
		//Go go setting
		solo.pressMenuItem(2);
		Log.e("Log 1","----------->GO TO SETTING");
		//Go to reset data
		solo.sendKey(20);
		solo.sendKey(20);
		solo.sendKey(20);
		solo.sendKey(20);
		solo.sendKey(20);
		solo.sendKey(20);
		solo.sendKey(20);
		solo.sendKey(20);
		solo.sendKey(20);
		Log.e("Log 2","----------->GO TO RESET");
		solo.sendKey(66);
		solo.sleep(2000);
		Log.e("----------->CHECK TO RESET", "Log 3");
		Log.e("----------->Message: "+solo.getString(R.string.msgAreYouSureToResetData), "Log 3");
		//confirmed reset data
		solo.clickOnButton(0);
		Log.e("Log 4", "----------->CONFIRMED");
		//confirmed logging out
		solo.sleep(5000);
		boolean expected2 = true;
		boolean actual2 = solo.searchText("Already a member at Doximity.com?");
		assertEquals("You are now LOGGED OUT", expected2, actual2);
	}*/
	
	//******************************************************************************************
	//END DEVELOPING TESTCASE
	//******************************************************************************************

	@Override
	public void	tearDown()throws Exception
	{
		try
			{
				solo.finalize();
			}
		catch (Throwable e)
			{
				e.printStackTrace();
			}
		getActivity().finish();
		super.tearDown();
	}	
}
