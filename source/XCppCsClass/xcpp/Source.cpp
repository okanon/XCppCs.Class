#include <iostream>
#include <cstdlib>
#include <comutil.h>
#include <objbase.h>
#include <tchar.h>

#pragma comment(lib, "comsupp.lib")
#pragma comment(lib, "comsuppw.lib")

IDispatch *pidsp = NULL;
IUnknown *piunk = NULL;

//long _Init(void);
long _XInit(std::string classes);
long _Finalize(void);
long _Sum(long num1, long num2);
long _Multi(long num1, long num2);
long _Once(long num);
//BSTR _Name(std::string name, long age);
std::string _Name(std::string name, long age);
void _Hello(void);

int _tmain(int argc, _TCHAR* argv[]) {
	int r = 0;
	BSTR brs;
	std::string rs;

	_XInit("xcs.XCsClass");

	r = _Sum(100, 200);
	std::cout << "Sum=" << r << std::endl;

	r = _Multi(200, 400);
	std::cout << "Multi=" << r << std::endl;

	rs = _Name("Okanon", 20);
	std::cout << rs.c_str() << std::endl;
	/*brs = _Name("Okanon", 20);
	printf("%S\n", brs);*/

	r = _Once(37);
	std::cout << "Once=" << r << std::endl;

	_Hello();

	_Finalize();

	system("PAUSE");
}

std::wstring str_to_wchar(std::string arg) {
	std::string narrow_string(arg);
	std::wstring wide_string = std::wstring(narrow_string.begin(), narrow_string.end());
	return wide_string;
}

std::string& bstr_to_str(const BSTR bstr, std::string& dst, int cp = CP_UTF8)
{
	if (!bstr)
	{
		dst.clear();
		return dst;
	}

	int res = WideCharToMultiByte(cp, 0, bstr, -1, NULL, 0, NULL, NULL);
	if (res > 0)
	{
		dst.resize(res);
		WideCharToMultiByte(cp, 0, bstr, -1, &dst[0], res, NULL, NULL);
	}
	else
	{
		dst.clear();
	}
	return dst;
}

std::string bstr_to_str(BSTR bstr, int cp = CP_UTF8)
{
	std::string str;
	bstr_to_str(bstr, str, cp);
	return str;
}

/*
long _Init(void) {
	CLSID clsid;

	::CoInitialize(0);

	try {
		HRESULT h_result = CLSIDFromProgID(L"xcs.XCsClass", &clsid);
		if (FAILED(h_result)) {
			throw std::exception("DLL import Exception");
		}

		h_result = CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER, IID_IUnknown, (void**)&piunk);
		if (FAILED(h_result)) {
			throw std::exception("Create Instance Exception");
		}

		h_result = piunk->QueryInterface(IID_IDispatch, (void**)&pidsp);
		if (FAILED(h_result)) {
			throw std::exception("Get Instance Exception");
		}
	}
	catch (std::exception &ex) {
		std::cerr << ex.what() << std::endl;
	}

	return 0;
}*/

long _XInit(std::string classes) {
	CLSID clsid;

	::CoInitialize(0);

	try {
		HRESULT h_result = CLSIDFromProgID(str_to_wchar(classes).c_str(), &clsid);
		if (FAILED(h_result)) {
			throw std::exception("DLL import Exception");
		}

		h_result = CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER, IID_IUnknown, (void**)&piunk);
		if (FAILED(h_result)) {
			throw std::exception("Create Instance Exception");
		}

		h_result = piunk->QueryInterface(IID_IDispatch, (void**)&pidsp);
		if (FAILED(h_result)) {
			throw std::exception("Get Instance Exception");
		}
	}
	catch (std::exception &ex) {
		std::cerr << ex.what() << std::endl;
	}

	return 0;
}

long _Finalize(void) {
	pidsp->Release();
	piunk->Release();
	::CoUninitialize();

	return 0;
}

long _Sum(long num1, long num2) {
	DISPID disp = 0;
	LPOLESTR func = L"Sum";
	HRESULT h_result = pidsp->GetIDsOfNames(IID_NULL, &func, 1, LOCALE_SYSTEM_DEFAULT, &disp);
	if (FAILED(h_result)) {
		throw std::exception("Get IDs Of Names failed.");
	}

	DISPPARAMS params;
	::memset(&params, 0, sizeof(DISPPARAMS));

	params.cNamedArgs = 0;
	params.rgdispidNamedArgs = NULL;
	params.cArgs = 2;

	VARIANTARG* pval = new VARIANTARG[params.cArgs];

	pval[0].vt = VT_I4;
	pval[0].lVal = num2;
	pval[1].vt = VT_I4;
	pval[1].lVal = num1;
	params.rgvarg = pval;

	VARIANT vret;
	VariantInit(&vret);

	pidsp->Invoke(disp, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, &params, &vret, NULL, NULL);

	delete[] pval;
	return vret.iVal;
}

long _Multi(long num1, long num2) {
	DISPID disp = 0;
	LPOLESTR func = L"Multi";
	HRESULT h_result = pidsp->GetIDsOfNames(IID_NULL, &func, 1, LOCALE_SYSTEM_DEFAULT, &disp);
	if (FAILED(h_result)) {
		throw std::exception("Get IDs Of Names failed.");
	}

	DISPPARAMS params;
	::memset(&params, 0, sizeof(DISPPARAMS));

	params.cNamedArgs = 0;
	params.rgdispidNamedArgs = NULL;
	params.cArgs = 2;

	VARIANTARG* pval = new VARIANTARG[params.cArgs];

	pval[0].vt = VT_I4;
	pval[0].lVal = num2;
	pval[1].vt = VT_I4;
	pval[1].lVal = num1;
	params.rgvarg = pval;

	VARIANT vret;
	VariantInit(&vret);

	pidsp->Invoke(disp, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, &params, &vret, NULL, NULL);

	delete[] pval;
	return vret.lVal;
}

//BSTR _Name ....
std::string _Name(std::string name, long age) {
	DISPID disp = 0;
	LPOLESTR func = L"Name";
	HRESULT h_result = pidsp->GetIDsOfNames(IID_NULL, &func, 1, LOCALE_SYSTEM_DEFAULT, &disp);
	if (FAILED(h_result)) {
		throw std::exception("Get IDs Of Names failed.");
	}

	DISPPARAMS params;
	::memset(&params, 0, sizeof(DISPPARAMS));

	params.cNamedArgs = 0;
	params.rgdispidNamedArgs = NULL;
	params.cArgs = 2;

	VARIANTARG* pval = new VARIANTARG[params.cArgs];

	pval[0].vt = VT_I4;
	pval[0].lVal = age;
	pval[1].vt = VT_BSTR;
	pval[1].bstrVal = SysAllocString(str_to_wchar(name).c_str());
	params.rgvarg = pval;

	VARIANT vret;
	VariantInit(&vret);

	pidsp->Invoke(disp, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, &params, &vret, NULL, NULL);

	delete[] pval;

	//return vret.bstrVal;
	return bstr_to_str(vret.bstrVal);
}

long _Once(long num) {
	DISPID disp = 0;
	LPOLESTR func = L"Sum";
	HRESULT h_result = pidsp->GetIDsOfNames(IID_NULL, &func, 1, LOCALE_SYSTEM_DEFAULT, &disp);
	if (FAILED(h_result)) {
		throw std::exception("Get IDs Of Names failed.");
	}

	DISPPARAMS params;
	::memset(&params, 0, sizeof(DISPPARAMS));

	params.cNamedArgs = 0;
	params.rgdispidNamedArgs = NULL;
	params.cArgs = 2;

	VARIANTARG* pval = new VARIANTARG[params.cArgs];

	pval->vt = VT_I4;
	pval->lVal = num;
	params.rgvarg = pval;

	VARIANT vret;
	VariantInit(&vret);

	pidsp->Invoke(disp, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, &params, &vret, NULL, NULL);

	delete[] pval;
	return vret.lVal;
}

void _Hello(void) {
	DISPID disp = 0;
	LPOLESTR func = L"Hello";
	HRESULT h_result = pidsp->GetIDsOfNames(IID_NULL, &func, 1, LOCALE_SYSTEM_DEFAULT, &disp);
	if (FAILED(h_result)) {
		throw std::exception("Get IDs Of Names failed.");
	}

	DISPPARAMS params;
	::memset(&params, 0, sizeof(DISPPARAMS));

	VARIANT vret;
	VariantInit(&vret);

	pidsp->Invoke(disp, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, &params, &vret, NULL, NULL);
}

