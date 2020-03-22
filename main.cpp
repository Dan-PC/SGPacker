#ifndef UNICODE
#define UNICODE
#endif
#define WIN32_LEAN_AND_MEAN
#define VC_EXTRALEAN

#include <windows.h>
#include <Shobjidl.h>
#include <fstream>

void Assert(HRESULT hr) {
	if (FAILED(hr)) {
		if (IsProcessorFeaturePresent(PF_FASTFAIL_AVAILABLE)) { __fastfail(0); }
		terminate();
	}
}

IShellItemArray* GetInputFiles() {
	IFileOpenDialog* pFileOpen;

	HRESULT hr = CoCreateInstance(CLSID_FileOpenDialog, 0, CLSCTX_ALL, __uuidof(pFileOpen), reinterpret_cast<void**>(&pFileOpen));
	Assert(hr);

	pFileOpen->SetTitle(L"Select the files to pack.");

	FILEOPENDIALOGOPTIONS opt = 0;
	pFileOpen->GetOptions(&opt);
	hr = pFileOpen->SetOptions(opt | FOS_ALLOWMULTISELECT);
	Assert(hr);

	hr = pFileOpen->Show(0);
	Assert(hr);

	IShellItemArray* pItemArray;
	hr = pFileOpen->GetResults(&pItemArray);
	Assert(hr);

	pFileOpen->Release();
	return pItemArray;
}

IShellItem* GetOutputFile() {
	IFileSaveDialog* pFileSave;
	IShellItem* ret = 0;

	HRESULT hr = CoCreateInstance(CLSID_FileSaveDialog, 0, CLSCTX_ALL, __uuidof(pFileSave), reinterpret_cast<void**>(&pFileSave));
	Assert(hr);

	pFileSave->SetTitle(L"Select the output file.");

	FILEOPENDIALOGOPTIONS fos = 0;
	pFileSave->GetOptions(&fos);
	pFileSave->SetOptions(fos | FOS_CREATEPROMPT | FOS_STRICTFILETYPES | FOS_FORCEFILESYSTEM);

	COMDLG_FILTERSPEC rgSpec = { L"mpk", L"*.mpk" };
	pFileSave->SetFileTypes(1, &rgSpec);
	pFileSave->SetDefaultExtension(rgSpec.pszName);

	hr = pFileSave->Show(0);
	Assert(hr);

	IShellItem* pItem;
	hr = pFileSave->GetResult(&pItem);
	Assert(hr);

	pFileSave->Release();
	return pItem;
}

const char header[8]{ 0x4d, 0x50, 0x4b, 0000, 0000, 0000, 0002, 0000 };


int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE, PWSTR pCmdLine, int nCmdShow)
{
	auto hr = CoInitializeEx(0, COINIT_MULTITHREADED | COINIT_DISABLE_OLE1DDE);
	Assert(hr);

	auto pInput = GetInputFiles();
	auto pOut = GetOutputFile();

	DWORD dwNumItems = 0;
	hr = pInput->GetCount(&dwNumItems);
	Assert(hr);

	LPWSTR pszName = 0;
	hr = pOut->GetDisplayName(SIGDN_FILESYSPATH, &pszName);
	Assert(hr);

	auto output = std::fstream(pszName, std::ios::trunc | std::ios::binary | std::ios::in | std::ios::out);
	if (output.is_open()) {
		output.write(header, 8);
		output.write(reinterpret_cast<char*>(&dwNumItems), 4);



		std::ifstream input;

		size_t fileblock_pos = (0x40 + dwNumItems * 256);		//position of the first entry file in the archive, each entry is 256 bytes long
		fileblock_pos += (0x800 - (fileblock_pos % 0x800)) & 0x7ff;					//must be a multiple of 0x800
		//unsigned skipped = 0;
		for (size_t i = 0; i < dwNumItems; i++) {
			IShellItem* psi;
			hr = pInput->GetItemAt(i, &psi);
			Assert(hr);
			LPWSTR pszInput;
			psi->GetDisplayName(SIGDN_FILESYSPATH, &pszInput);
			wchar_t name[224]{ 0 };
			wchar_t ext[16]{ 0 };
			_wsplitpath_s(pszInput, nullptr, 0, nullptr, 0, name, 224, ext, 16);
			wcscat_s(name, ext);

			char name_ansi[224];
			WideCharToMultiByte(CP_ACP, WC_NO_BEST_FIT_CHARS, name, 224, name_ansi, 224, NULL, NULL);

			if (*reinterpret_cast<DWORD*>(name_ansi) == 0x5f4e4957) { //if it starts with "WIN_", change it to "WIN\"
				name_ansi[3] = 0x5c;
			}
			else if ((*reinterpret_cast<DWORD*>(name_ansi) & 0xffffff) == 0x5f4433) {	//if it starts with "3D_", change it to "3D\"
				name_ansi[2] = 0x5c;
			}
			input.open(pszInput, std::ios::in | std::ios::binary | std::ios::ate);
			if (!input.is_open()) {
				CoTaskMemFree(pszInput);
				psi->Release();
				break;
			}
			size_t inputsize = input.tellg();
			if (!inputsize) {
				MessageBox(0, name, L"File corrupted.", 0);

				input.close();
				CoTaskMemFree(pszInput);
				psi->Release();
				break;
			}
			input.seekg(0);

			output.seekp(0x44 + i * 256);
			output.write(reinterpret_cast<char*>(&i), 4);				//entry number
			output.write(reinterpret_cast<char*>(&fileblock_pos), 8);	//offset of the packed file from the beginning of the archive
			output.write(reinterpret_cast<char*>(&inputsize), 8);		//write file size twice
			output.write(reinterpret_cast<char*>(&inputsize), 8);
			output << name_ansi;										//entry name, 224 bytes

			output.seekp(fileblock_pos);
			output << input.rdbuf();		//write the actual file

			if (i + 1 < dwNumItems) {
				fileblock_pos = output.tellp();
				fileblock_pos += (0x800 - (fileblock_pos % 0x800)) & 0x7ff;		//offset for writing the next file
			}
			input.close();
			CoTaskMemFree(pszInput);
			psi->Release();
		}
		output.close();
	}
	CoTaskMemFree(pszName);
	pOut->Release();
	pInput->Release();
	CoUninitialize();
	return 0;
}

