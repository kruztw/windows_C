#include <windows.h>
#include <objbase.h>
#include "IExample.h"

static DWORD		OutstandingObjects;
static DWORD		LockCount;

typedef struct {
	IExampleVtbl	*lpVtbl;
	DWORD			count;
	char			buffer[80];
} MyRealIExample;


static HRESULT STDMETHODCALLTYPE QueryInterface(IExample *this, REFIID vTableGuid, void **ppv)
{
	if (!IsEqualIID(vTableGuid, &IID_IUnknown) && !IsEqualIID(vTableGuid, &IID_IExample))
	{
      *ppv = 0;
      return(E_NOINTERFACE);
	}

	*ppv = this;
	this->lpVtbl->AddRef(this);
	return(NOERROR);
}

static ULONG STDMETHODCALLTYPE AddRef(IExample *this)
{
	return(++((MyRealIExample *)this)->count);
}

static ULONG STDMETHODCALLTYPE Release(IExample *this)
{

	if (--((MyRealIExample *)this)->count == 0)
	{
		GlobalFree(this);
		InterlockedDecrement(&OutstandingObjects);
		return(0);
	}
	return(((MyRealIExample *)this)->count);
}

static HRESULT STDMETHODCALLTYPE SetString(IExample *this, char *str)
{
	DWORD  i;

	if (!str) return(E_POINTER);

	i = lstrlen(str);
	if (i > 79) i = 79;
	CopyMemory(((MyRealIExample *)this)->buffer, str, i);
	((MyRealIExample *)this)->buffer[i] = 0;

	return(NOERROR);
}

static HRESULT STDMETHODCALLTYPE GetString(IExample *this, char *buffer, DWORD length)
{
	DWORD  i;

	if (!buffer) return(E_POINTER);

	if (length)
	{
		i = lstrlen(((MyRealIExample *)this)->buffer);
		--length;
		if (i > length) i = length;
		CopyMemory(buffer, ((MyRealIExample *)this)->buffer, i);
		buffer[i] = 0;
	}
	return(NOERROR);
}

static const IExampleVtbl IExample_Vtbl = 
{
    QueryInterface,
    AddRef,
    Release,
    SetString,
    GetString
};


static IClassFactory	MyIClassFactoryObj;

static ULONG STDMETHODCALLTYPE classAddRef(IClassFactory *this)
{
	InterlockedIncrement(&OutstandingObjects);
	return(1);
}

static HRESULT STDMETHODCALLTYPE classQueryInterface(IClassFactory *this, REFIID factoryGuid, void **ppv)
{
	if (IsEqualIID(factoryGuid, &IID_IUnknown) || IsEqualIID(factoryGuid, &IID_IClassFactory))
	{
		this->lpVtbl->AddRef(this);
		*ppv = this;
		return(NOERROR);
	}

	*ppv = 0;
	return(E_NOINTERFACE);
}

static ULONG STDMETHODCALLTYPE classRelease(IClassFactory *this)
{
	return(InterlockedDecrement(&OutstandingObjects));
}

static HRESULT STDMETHODCALLTYPE classCreateInstance(IClassFactory *this, IUnknown *punkOuter, REFIID vTableGuid, void **objHandle)
{
	HRESULT				hr;
	register IExample	*thisobj;

	*objHandle = 0;
	if (punkOuter)
		hr = CLASS_E_NOAGGREGATION;
	else
	{
		if (!(thisobj = (IExample *)GlobalAlloc(GMEM_FIXED, sizeof(MyRealIExample))))
			hr = E_OUTOFMEMORY;
		else
		{
			thisobj->lpVtbl = (IExampleVtbl *)&IExample_Vtbl;
			((MyRealIExample *)thisobj)->count = 1;
			((MyRealIExample *)thisobj)->buffer[0] = 0;
			hr = IExample_Vtbl.QueryInterface(thisobj, vTableGuid, objHandle);
			IExample_Vtbl.Release(thisobj);

			if (!hr) InterlockedIncrement(&OutstandingObjects);
		}
	}
	return(hr);
}

static HRESULT STDMETHODCALLTYPE classLockServer(IClassFactory *this, BOOL flock)
{
	if (flock) InterlockedIncrement(&LockCount);
	else InterlockedDecrement(&LockCount);

	return(NOERROR);
}

static const IClassFactoryVtbl IClassFactory_Vtbl = 
{
	classQueryInterface,
    classAddRef,
    classRelease,
    classCreateInstance,
    classLockServer
};


HRESULT PASCAL DllGetClassObject(REFCLSID objGuid, REFIID factoryGuid, void **factoryHandle)
{
	register HRESULT		hr;

	if (IsEqualCLSID(objGuid, &CLSID_IExample))
	{
		hr = classQueryInterface(&MyIClassFactoryObj, factoryGuid, factoryHandle);
	}
	else
	{
		*factoryHandle = 0;
		hr = CLASS_E_CLASSNOTAVAILABLE;
	}

	return(hr);
}


HRESULT PASCAL DllCanUnloadNow(void)
{
	return((OutstandingObjects | LockCount) ? S_FALSE : S_OK);
}

BOOL WINAPI DllMain(HINSTANCE instance, DWORD fdwReason, LPVOID lpvReserved)
{
	switch (fdwReason)
	{
		case DLL_PROCESS_ATTACH:
		{
			OutstandingObjects = LockCount = 0;
			MyIClassFactoryObj.lpVtbl = (IClassFactoryVtbl *)&IClassFactory_Vtbl;
			DisableThreadLibraryCalls(instance);
		}
	}

	return(1);
}