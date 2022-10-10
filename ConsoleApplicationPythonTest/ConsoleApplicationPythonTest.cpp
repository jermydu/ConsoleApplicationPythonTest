#include <Python.h>
#include<iostream>

using namespace std;

int main()
{
	//添加python.exe所在路径
	Py_SetPythonHome(L"C:\\Users\\dujinwei\\AppData\\Local\\Programs\\Python\\Python310");
	
	//python初始化
	Py_Initialize();
	if (!Py_IsInitialized())
	{
		printf("初始化失败！");
		return 0;
	}
	else {
		PyRun_SimpleString("import sys");
		//修改Python路径
		PyRun_SimpleString("sys.path.append('./')");

		PyObject* pModule = NULL;
		PyObject* pFunc = NULL;

		//hello 测试
		//pModule = PyImport_ImportModule("hello");//这里是要调用的文件名hello.py
		//if (pModule == NULL)
		//{
		//	cout << "没找到该Python文件" << endl;
		//}
		//else {
		//	pFunc = PyObject_GetAttrString(pModule, "add");//这里是要调用的函数名
		//	PyObject* args = Py_BuildValue("(ii)", 28, 103);//给python函数参数赋值

		//	PyObject* pRet = PyObject_CallObject(pFunc, args);//调用函数

		//	int res = 0;
		//	PyArg_Parse(pRet, "i", &res);//转换返回类型

		//	cout << "res:" << res << endl;//输出结果
		//}

		pModule = PyImport_ImportModule("convert");
		if (pModule == NULL)
		{
			cout << "没找到convert.py文件" << endl;
		}
		else {
			pFunc = PyObject_GetAttrString(pModule, "convertfile2pdf");//这里是要调用的函数名
			
			std::string file_path = "C:\\Users\\dujinwei\\source\\repos\\ConsoleApplicationPythonTest\\x64\\Release\\test.xlsx";
			std::string pdf_path = "C:\\Users\\dujinwei\\source\\repos\\ConsoleApplicationPythonTest\\x64\\Release\\\\test.pdf";
			std::string png_path = "C:\\Users\\dujinwei\\source\\repos\\ConsoleApplicationPythonTest\\x64\\Release\\\\test.png";
			std::string file_type = "excel";
			
			PyObject* args = Py_BuildValue("(s,s,s)", file_path.c_str(), pdf_path.c_str(),file_type.c_str());

			PyObject* pRet = PyObject_CallObject(pFunc, args);

			int res = 0;
			//转换返回类型
			PyArg_Parse(pRet, "i", &res);

			cout << "res:" << res << endl;
			//这里是要调用的函数名
			pFunc = PyObject_GetAttrString(pModule, "pdf_image");

			PyObject* args2 = Py_BuildValue("(s,s,i,i,i)", pdf_path.c_str(),png_path.c_str(), 5, 5,0);

			//调用函数
			PyObject_CallObject(pFunc, args2);
		}

		//调用Py_Finalize，这个根Py_Initialize相对应的。
		Py_Finalize();
	}
	return 0;
}