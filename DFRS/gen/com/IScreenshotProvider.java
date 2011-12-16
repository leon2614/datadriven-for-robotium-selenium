/*
 * This file is auto-generated.  DO NOT MODIFY.
 * Original file: D:\\Projects\\Robotium\\DFRS\\src\\com\\IScreenshotProvider.aidl
 */
package com;
// Interface for fetching screenshots

public interface IScreenshotProvider extends android.os.IInterface
{
/** Local-side IPC implementation stub class. */
public static abstract class Stub extends android.os.Binder implements com.IScreenshotProvider
{
private static final java.lang.String DESCRIPTOR = "com.IScreenshotProvider";
/** Construct the stub at attach it to the interface. */
public Stub()
{
this.attachInterface(this, DESCRIPTOR);
}
/**
 * Cast an IBinder object into an com.IScreenshotProvider interface,
 * generating a proxy if needed.
 */
public static com.IScreenshotProvider asInterface(android.os.IBinder obj)
{
if ((obj==null)) {
return null;
}
android.os.IInterface iin = (android.os.IInterface)obj.queryLocalInterface(DESCRIPTOR);
if (((iin!=null)&&(iin instanceof com.IScreenshotProvider))) {
return ((com.IScreenshotProvider)iin);
}
return new com.IScreenshotProvider.Stub.Proxy(obj);
}
public android.os.IBinder asBinder()
{
return this;
}
@Override public boolean onTransact(int code, android.os.Parcel data, android.os.Parcel reply, int flags) throws android.os.RemoteException
{
switch (code)
{
case INTERFACE_TRANSACTION:
{
reply.writeString(DESCRIPTOR);
return true;
}
case TRANSACTION_isAvailable:
{
data.enforceInterface(DESCRIPTOR);
boolean _result = this.isAvailable();
reply.writeNoException();
reply.writeInt(((_result)?(1):(0)));
return true;
}
case TRANSACTION_takeScreenshot:
{
data.enforceInterface(DESCRIPTOR);
java.lang.String _result = this.takeScreenshot();
reply.writeNoException();
reply.writeString(_result);
return true;
}
}
return super.onTransact(code, data, reply, flags);
}
private static class Proxy implements com.IScreenshotProvider
{
private android.os.IBinder mRemote;
Proxy(android.os.IBinder remote)
{
mRemote = remote;
}
public android.os.IBinder asBinder()
{
return mRemote;
}
public java.lang.String getInterfaceDescriptor()
{
return DESCRIPTOR;
}
// Checks whether the native background application is running
// (and thus whether the screenshots are available)

public boolean isAvailable() throws android.os.RemoteException
{
android.os.Parcel _data = android.os.Parcel.obtain();
android.os.Parcel _reply = android.os.Parcel.obtain();
boolean _result;
try {
_data.writeInterfaceToken(DESCRIPTOR);
mRemote.transact(Stub.TRANSACTION_isAvailable, _data, _reply, 0);
_reply.readException();
_result = (0!=_reply.readInt());
}
finally {
_reply.recycle();
_data.recycle();
}
return _result;
}
// Create a screen snapshot and returns path to file where it is written.

public java.lang.String takeScreenshot() throws android.os.RemoteException
{
android.os.Parcel _data = android.os.Parcel.obtain();
android.os.Parcel _reply = android.os.Parcel.obtain();
java.lang.String _result;
try {
_data.writeInterfaceToken(DESCRIPTOR);
mRemote.transact(Stub.TRANSACTION_takeScreenshot, _data, _reply, 0);
_reply.readException();
_result = _reply.readString();
}
finally {
_reply.recycle();
_data.recycle();
}
return _result;
}
}
static final int TRANSACTION_isAvailable = (android.os.IBinder.FIRST_CALL_TRANSACTION + 0);
static final int TRANSACTION_takeScreenshot = (android.os.IBinder.FIRST_CALL_TRANSACTION + 1);
}
// Checks whether the native background application is running
// (and thus whether the screenshots are available)

public boolean isAvailable() throws android.os.RemoteException;
// Create a screen snapshot and returns path to file where it is written.

public java.lang.String takeScreenshot() throws android.os.RemoteException;
}
