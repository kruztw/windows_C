#include <windows.h>
#include <objbase.h>
#include <stdio.h>
#include "../IExample/IExample.h"

int main(int argc, char **argv)
{
	IExample		*exampleObj;
	IClassFactory	*classFactory;
	HRESULT			hr;

	if (!CoInitialize(0))
	{
		if ((hr = CoCreateInstance(&CLSID_IExample, 0, CLSCTX_INPROC_SERVER, &IID_IExample, (LPVOID*)&exampleObj)))
			MessageBox(0, "Can't create IExample object", "CoCreateInstance error", MB_OK | MB_ICONEXCLAMATION);
		/*if ((hr = CoGetClassObject(&CLSID_IExample, CLSCTX_INPROC_SERVER, 0, &IID_IClassFactory, &classFactory)))
			MessageBox(0, "Can't get IClassFactory", "CoGetClassObject error", MB_OK|MB_ICONEXCLAMATION);*/
		else
		{
			/*if ((hr = classFactory->lpVtbl->CreateInstance(classFactory, 0, &IID_IExample, &exampleObj)))
			{
				classFactory->lpVtbl->Release(classFactory);
				MessageBox(0, "Can't create IExample object", "CreateInstance error", MB_OK|MB_ICONEXCLAMATION);
			}
			else*/
			{
				char	buffer[80];
				//classFactory->lpVtbl->Release(classFactory);
				exampleObj->lpVtbl->SetString(exampleObj, "Hello world!");
				exampleObj->lpVtbl->GetString(exampleObj, buffer, sizeof(buffer));
				printf("string = %s (should be Hello World!)\n", buffer);
				exampleObj->lpVtbl->Release(exampleObj);
			}
		}

		CoUninitialize();
	}
	else
		MessageBox(0, "Can't initialize COM", "CoInitialize error", MB_OK|MB_ICONEXCLAMATION);

	return(0);
}