package com.dfrs;

import java.io.File;
import java.io.FileOutputStream;

import android.graphics.Bitmap;
import android.os.Environment;
import android.util.Log;
import android.view.View;

public class Screenshot {

private static final String SCREEN_SHOTS_LOCATION = "mnt/sdcard/DCIM/DFRS";

public static void takeScreenShot(View view, String name, String project_folder) throws Exception {
  view.setDrawingCacheEnabled(true);
  view.buildDrawingCache();
  Bitmap b = view.getDrawingCache();
  FileOutputStream fos = null;
  try {
   File SDCardRoot = Environment.getExternalStorageDirectory();
   File sddir = new File(SDCardRoot.getAbsolutePath() + "/DCIM/DFRS/"+project_folder);
   //File sddir = new File(Environment.getExternalStorageDirectory().getPath() + "/screenshots/");
   if (!sddir.exists()) {
   sddir.mkdirs();
   }
   fos = new FileOutputStream(sddir.getPath() + "/" +  name + ".png");
   if (fos != null) {
   b.compress(Bitmap.CompressFormat.JPEG, 90, fos);
   fos.close();
   }
  } catch (Exception e) {
   
  }
}
}