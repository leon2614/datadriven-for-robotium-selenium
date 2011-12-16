package com.dfrs;

import android.test.ActivityInstrumentationTestCase2;
import android.test.suitebuilder.annotation.Smoke;
import android.util.Log;
import com.IScreenshotProvider;
import com.ScreenshotService;

import android.content.Intent;
import android.app.Activity;
import android.content.ServiceConnection;
import android.os.IBinder;
import android.content.ComponentName;
import android.content.Context;

import com.jayway.android.robotium.solo.Solo;

@SuppressWarnings("unchecked")
public class RobotiumTest extends ActivityInstrumentationTestCase2
{
	private static final String	TARGET_PACKAGE_ID = "";
	private static final String LAUNCHER_ACTIVITY_FULL_CLASSNAME = ".app.screens.GenSplash";
	private	static Class<?>	launcherActivityClass;
	
	//ASL screenshot provider
	private ServiceConnection aslServiceConn = new ServiceConnection() {

		@Override
		public void onServiceDisconnected(ComponentName name) {
			// TODO Auto-generated method stub

		}

		@Override
		public void onServiceConnected(ComponentName name, IBinder service) {
			Log.e("KoolJ_ASL", "Started!");
			aslProvider = IScreenshotProvider.Stub.asInterface(service);
			
		}
	};
	public static IScreenshotProvider aslProvider = null;
	
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
		}
	public Solo solo;

	@Override 
	protected void setUp() throws Exception 
	{
		solo=new Solo(getInstrumentation(),	getActivity());
        Intent intent = new Intent();
        intent.setClass(getActivity(), ScreenshotService.class);
        //intent.addCategory(Intent.ACTION_DEFAULT);
        getActivity().bindService (intent, aslServiceConn, Context.BIND_AUTO_CREATE);
	}
	
	//******************************************************************************************
	//START DEVELOPING TESTCASE
	//******************************************************************************************

	@Smoke
	public void test0run()
	{
		
		//Call config CSV file
		KoolJ_datadriven Kjdriven=new KoolJ_datadriven();
		Kjdriven.openconfig("/config.xls","/output.xls", solo);
	}

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
