/*
2017-03-13 dbj@dbj.org ongoing transformation into modern C++
*/

////////////////////////////////////////////////////////////////////////////////
//
// Testing WinReg
//
// by Giovanni Dicanio <giovanni.dicanio@gmail.com>
//
////////////////////////////////////////////////////////////////////////////////
#include <iostream>
#include "wreg.h"   // WinReg public header

using std::wcout;
using std::wstring;
using std::vector;

// Helpers
wstring ToHexString(BYTE b);
wstring ToHexString(DWORD dw);
void PrintRegValue(const winreg::RegValue& value);

winreg::RegKey return_key_with_some_values (const std::wstring & testKeyName)
{
	wcout << L"Creating some test key and writing some values into it...\n";

	winreg::RegKey key = winreg::RegKey::CreateKey(HKEY_CURRENT_USER, testKeyName);

	winreg::RegValue v(L"Test DWORD", REG_DWORD);

	v.Dword() = 0x64;
	key.SetValue(v);

	v.Reset(REG_SZ, L"Test REG_SZ");
	v.String() = L"Hello World";
	key.SetValue(v);

	v.Reset(REG_EXPAND_SZ,L"Test REG_EXPAND_SZ");
	v.ExpandString() = L"%WinDir%";
	key.SetValue(v);

	v.Reset(REG_MULTI_SZ,L"Test REG_MULTI_SZ");
	v.MultiString().push_back(L"Ciao");
	v.MultiString().push_back(L"Hi");
	v.MultiString().push_back(L"Connie");
	key.SetValue(v);

	v.Reset(REG_BINARY,L"Test REG_BINARY");
	v.Binary().push_back(0x22);
	v.Binary().push_back(0x33);
	v.Binary().push_back(0x44);
	key.SetValue(v);

	return key;
}

//
// Test Delete
//
void test_delete(const std::wstring & testKeyName )
{
	wcout << L"\nDeleting a value...\n";

	winreg::RegKey key = winreg::RegKey::OpenKey(HKEY_CURRENT_USER, testKeyName, KEY_WRITE | KEY_READ);

	// Delete a value
	const wstring valueName = L"TestValue_DWORD";
	winreg::DeleteValue(key.Handle(), valueName);

	// Try accessing a non-existent value
	try
	{
		wcout << "Trying accessing value just deleted...\n";
		winreg::RegValue value = winreg::QueryValue(key.Handle(), valueName);
	}
	catch (const winreg::RegException& ex)
	{
		wcout << L"winreg::RegException correctly caught.\n";
		wcout << L"Error code: " << ex.ErrorCode() << L'\n';
		// Error code should be 2: ERROR_FILE_NOT_FOUND
		if (ex.ErrorCode() == ERROR_FILE_NOT_FOUND)
		{
			wcout << L"All right, I expected ERROR_FILE_NOT_FOUND (== 2).\n\n";
		}
	}

	// Delete the whole key --> from the REGISTRY that is!
	winreg::DeleteKey(HKEY_CURRENT_USER, testKeyName);
	// after this destructor release the key object in memory
}

//
// Enum Values
//
void enum_values(winreg::RegKey & key)
{
	const vector<wstring> valueNames = winreg::EnumerateValueNames(key.Handle());

	for (const auto& valueName : valueNames)
	{
		winreg::RegValue value = winreg::QueryValue(key.Handle(), valueName);
		wcout << valueName
			<< L" of type: " << winreg::ValueTypeIdToString(value.GetType())
			<< L"\n";

		PrintRegValue(value);
		wcout << L"\n";
	}
}

/*
*/
int main()
{
    wcout << "*** Testing WinReg -- by Giovanni Dicanio & DBJDBJ ***\n\n";

    const wstring testKeyName = 
		L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\FolderDescriptions\\";
	try {
		//
		// Enum Sub Keys
		//
		{
			wcout << L"\nEnumerating sub keys:\n";

			// KEY_READ is default access req. and this is bad to hide
			auto key = winreg::RegKey::OpenKey(HKEY_LOCAL_MACHINE, testKeyName);

			const vector<wstring> subNames = winreg::EnumerateSubKeyNames(key.Handle());

			for (const auto& subName : subNames)
			{
				winreg::RegKey subKey = winreg::RegKey::OpenKey(key.Handle(), subName);
				wcout << L"For the key:" << subName << L" Values are \n";
				enum_values(subKey);
				wcout << "-----------------------------------------------------------------\n";
			}
		}

	}
	catch ( winreg::RegException & rx ) {
		std::wcerr << L"winreg exception: " <<  rx.ErrorCode() << L" : " <<  rx.ErrorMessage() << std::endl ;
	}
	catch (... ) {
		std::wcerr << L"Unknown exception: " << std::endl;
	}
	// eof main
	return 0;
}


wstring ToHexString(BYTE b)
{
    wchar_t buf[10];
    swprintf_s(buf, L"0x%02X", b);
    return wstring(buf);
}


wstring ToHexString(DWORD dw)
{
    wchar_t buf[20];
    swprintf_s(buf, L"0x%08X", dw);
    return wstring(buf);
}


void PrintRegValue(const winreg::RegValue& value)
{
	wcout << L"Name:[" << value.name() << "]\tValue:\t";
    switch (value.GetType())
    {
    case REG_NONE:
    {
        wcout << L"None\t";
    }
    break;

    case REG_BINARY:
    {
        const vector<BYTE> & data = value.Binary();
        for (BYTE x : data)
        {
            wcout << ToHexString(x) << L" ";
        }
        wcout << L"\t";
    }
    break;

    case REG_DWORD:
    {
        DWORD dw = value.Dword();
        wcout << ToHexString(dw) << L'\t';
    }
    break;

    case REG_EXPAND_SZ:
    {
        wcout << L"[" << value.ExpandString() << L"]\t";

        wcout << L"Expanded: [" << winreg::ExpandEnvironmentStrings(value.ExpandString()) << "]\t";
    }
    break;

    case REG_MULTI_SZ:
    {
        const vector<wstring>& multiString = value.MultiString();
        for (const auto& s : multiString)
        {
            wcout << L"[" << s << L"]\t";
        }
    }
    break;

    case REG_SZ:
    {
        wcout << L"[" << value.String() << L"]\t";
    }
    break;

    default:
        wcout << L"Unsupported/Unknown registry value type\t";
        break;
    }
}