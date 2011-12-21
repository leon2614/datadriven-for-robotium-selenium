package com.dfrs;

import android.test.ActivityInstrumentationTestCase2;
import android.test.suitebuilder.annotation.Smoke;
import android.util.Log;

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
	}
	public Solo solo;

	@Override 
	protected void setUp() throws Exception 
	{
		solo=new Solo(getInstrumentation(),	getActivity());
	}

	@Smoke
	public void test0run()
	{
		
		//Call config CSV file
		KoolJ_datadriven Kjdriven=new KoolJ_datadriven();
		Kjdriven.openconfig("/config.xls","/output.xls", solo);
	}

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
