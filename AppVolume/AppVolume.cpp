#include <windows.h>
#include <string>
#include <Audiopolicy.h>
#include <Mmdeviceapi.h>
#include <Endpointvolume.h>
#include <Functiondiscoverykeys_devpkey.h>
#include <Shlwapi.h>
#include <Psapi.h>
#include <vector>
#include "..\..\API\RainmeterAPI.h"

#define SAFE_RELEASE(punk)  \
			  if ((punk) != nullptr) { (punk)->Release(); (punk) = nullptr; }
enum IndexType
{
	NOINDEXTYPE,
	NUMBER,
	APPNAME
};
enum NumberType
{
	NONUMTYPE,
	PEAK,
	VOLUME
};

enum StringType
{
	NOSTRTYPE,
	FILEPATH,
	FILENAME,
	ICONPATH
};

struct ChildMeasure;
struct AppSession;

struct ParentMeasure
{
	void* skin;
	void* rm;
	LPCWSTR name;
	BOOL initErr;
	ChildMeasure* ownerChild;
	BOOL ignoreSystemSound;
	std::vector<AppSession*> SessionCollection;
	ParentMeasure() :
		skin(nullptr),
		name(),
		ownerChild(nullptr),
		initErr(true),
		ignoreSystemSound(true)
	{
	}
};

struct ChildMeasure
{
	ParentMeasure* parent;
	IndexType indexType;
	NumberType numtype;
	StringType strtype;
	BOOL isParent;
	int indexNum;
	std::wstring indexApp;
	
	ChildMeasure() :
		indexType(NOINDEXTYPE),
		numtype(NONUMTYPE),
		strtype(NOSTRTYPE),
		isParent(false),
		indexNum(1),
		indexApp(L"")
	{
	}
};

struct AppSession
{
	IAudioMeterInformation * peak = nullptr;
	ISimpleAudioVolume * volume = nullptr;
	BOOL mute = false;
	DWORD processID = NULL;
	std::wstring appPath = L"";
	std::wstring appName = L"";
	
};

static BOOL com_initialized = FALSE;
const static CLSID CLSID_MMDeviceEnumerator = __uuidof(MMDeviceEnumerator);
const static IID IID_IMMDeviceEnumerator = __uuidof(IMMDeviceEnumerator);
const static IID IID_SessionManager2 = __uuidof(IAudioSessionManager2);
const static IID IID_SAV = __uuidof(ISimpleAudioVolume);
BOOL InitCom(ParentMeasure* measure);
BOOL UpdateList(ParentMeasure* measure);

IMMDeviceEnumerator * pEnumerator = nullptr;
std::wstring deviceName = L"";

std::vector<ParentMeasure*> g_ParentMeasures;


BOOL InitCom(ParentMeasure* measure)
{
	if (!com_initialized) com_initialized = SUCCEEDED(CoInitialize(0));
	if (!com_initialized)
	{
		RmLog(LOG_ERROR, L"AppVolume.dll: COM initialization failed");
		return false;
	}

	if (!SUCCEEDED(CoCreateInstance(CLSID_MMDeviceEnumerator, 0, CLSCTX_ALL, IID_IMMDeviceEnumerator, (void**)&pEnumerator)))
	{
		RmLog(LOG_ERROR, L"AppVolume.dll: COM creation failed");
		return false;
	}
	
	if (!SUCCEEDED(UpdateList(measure)))
		return false;

	CoUninitialize();
	com_initialized = FALSE;
	return true;
}

PLUGIN_EXPORT void Initialize(void** data, void* rm)
{
	ChildMeasure* child = new ChildMeasure;
	*data = child;

	void* skin = RmGetSkin(rm);

	std::wstring parentName = RmReadString(rm, L"Parent", L"");
	if (parentName.empty())
	{
		child->parent = new ParentMeasure;
		child->parent->name = RmGetMeasureName(rm);
		child->parent->skin = skin;
		child->parent->rm = rm;
		child->parent->ownerChild = child;
		child->parent->ignoreSystemSound = RmReadInt(rm, L"IgnoreSystemSound", 1) == 1;
		child->parent->initErr = !InitCom(child->parent);
		g_ParentMeasures.push_back(child->parent);
		child->isParent = true;
	}
	else
	{
		// Find parent using name AND the skin handle to be sure that it's the right one
		std::vector<ParentMeasure*>::const_iterator iter = g_ParentMeasures.begin();
		for (; iter != g_ParentMeasures.end(); ++iter)
		{
			if (_wcsicmp((*iter)->name, parentName.c_str()) == 0 &&
				(*iter)->skin == skin)
			{
				child->parent = (*iter);
				return;
			}
		}

		RmLog(LOG_ERROR, L"AppVolume.dll: Invalid Parent=");
	}
}

BOOL UpdateList(ParentMeasure* measure)
{
	IMMDevice * pDevice = nullptr;
	if (SUCCEEDED(pEnumerator->GetDefaultAudioEndpoint(eRender, eConsole, &pDevice)))
	{
		IAudioSessionManager2 * sessManager = nullptr;
		if (SUCCEEDED(pDevice->Activate(IID_SessionManager2, 7, 0, (void**)&sessManager)))
		{
			IPropertyStore * pProps = nullptr;
			if (SUCCEEDED(pDevice->OpenPropertyStore(STGM_READ, &pProps)))
			{
				PROPVARIANT varName;
				PropVariantInit(&varName);
				if (SUCCEEDED(pProps->GetValue(PKEY_Device_DeviceDesc, &varName)))
				{
					deviceName = varName.pwszVal;
				}
				PropVariantClear(&varName);
				SAFE_RELEASE(pProps);
			}
			SAFE_RELEASE(pDevice);
			IAudioSessionEnumerator * sessEnum = nullptr;
			if (SUCCEEDED(sessManager->GetSessionEnumerator(&sessEnum)))
			{
				SAFE_RELEASE(sessManager);
				std::vector<AppSession*>::const_iterator iter = measure->SessionCollection.begin();
				for (; iter != measure->SessionCollection.end(); ++iter)
				{
					SAFE_RELEASE((*iter)->volume);
					SAFE_RELEASE((*iter)->peak);
					delete *iter;
				}
				int sessionCount = 0;
				sessEnum->GetCount(&sessionCount);
				measure->SessionCollection.clear();
				measure->SessionCollection.shrink_to_fit();
				for (int i = 0; i < sessionCount; i++)
				{
					AppSession * newAppSession = new AppSession;

					IAudioSessionControl * sessControl = nullptr;
					if (!SUCCEEDED(sessEnum->GetSession(i, &sessControl)))
						continue;

					IAudioSessionControl2 * sessControl2 = nullptr;
					if (!SUCCEEDED(sessControl->QueryInterface(&sessControl2)))
					{
						SAFE_RELEASE(sessControl);
						continue;
					}
						
					LPWSTR sessInsID = L"";
					if (!SUCCEEDED(sessControl2->GetSessionInstanceIdentifier(&sessInsID)))
					{
						SAFE_RELEASE(sessControl);
						SAFE_RELEASE(sessControl2)
						continue;
					}

					std::wstring findWeirdChar = sessInsID;
					size_t weirdChar = findWeirdChar.find(L"#", 0);
					CoTaskMemFree(sessInsID);

					HRESULT isSystem = sessControl2->IsSystemSoundsSession();
					if ((measure->ignoreSystemSound && isSystem == S_OK) ||
						(weirdChar != std::wstring::npos && isSystem == S_FALSE))
					{
						SAFE_RELEASE(sessControl);
						SAFE_RELEASE(sessControl2);
						continue;
					}
					
					sessControl->QueryInterface(&newAppSession->volume);
					sessControl->QueryInterface(&newAppSession->peak);
					newAppSession->volume->GetMute(&newAppSession->mute);

					sessControl2->GetProcessId(&newAppSession->processID);

					SAFE_RELEASE(sessControl);
					SAFE_RELEASE(sessControl2);

					if (newAppSession->processID == 0)
					{
						newAppSession->appPath = newAppSession->appName = L"System Sound";
					}
					else
					{
						HANDLE checkProc = OpenProcess(PROCESS_ALL_ACCESS, false, newAppSession->processID);
						WCHAR procPath[400] = L"";
						if (checkProc != NULL)
						{
							GetModuleFileNameEx(checkProc, NULL, procPath, 400 / sizeof(WCHAR));
						}
						CloseHandle(checkProc);
						newAppSession->appPath = procPath;
						newAppSession->appName = PathFindFileName(newAppSession->appPath.c_str());
					}
					measure->SessionCollection.push_back(newAppSession);
				}
				SAFE_RELEASE(sessEnum);
			}
			else
			{
				RmLog(LOG_ERROR, L"AppVolume.dll: Could not enumerate AudioSessionManager");
				return false;
			}
		}
		else
		{
			RmLog(LOG_ERROR, L"AppVolume.dll: Could not activate AudioSessionManager");
			return false;
		}
	}
	else
	{
		RmLog(LOG_ERROR, L"AppVolume.dll: Could not get device");
		return false;
	}
	SAFE_RELEASE(pDevice);
	return true;
}

PLUGIN_EXPORT void Reload(void* data, void* rm, double* maxValue)
{

	ChildMeasure* child = (ChildMeasure*)data;
	ParentMeasure* parent = child->parent;
	if (!parent)
		return;

	if (child->isParent)
	{
		parent->ignoreSystemSound = RmReadInt(rm, L"IgnoreSystemSound", 1) == 1;
		return;
	}
	
	child->indexApp = RmReadString(rm, L"AppName", L"");
	if (!child->indexApp.empty())
	{
		child->indexType = APPNAME;
		child->indexNum = 0;
	}
	else
	{
		child->indexType = NUMBER;
		child->indexNum = RmReadInt(rm, L"Index", 1) - 1;
		child->indexApp = L"";
	}

	LPCWSTR type = RmReadString(rm, L"NumberType", L"VOLUME");
	if (_wcsicmp(type, L"VOLUME") == 0)
		child->numtype = VOLUME;
	else if (_wcsicmp(type, L"PEAK") == 0)
		child->numtype = PEAK;
	else
		RmLog(LOG_ERROR, L"AppVolume.dll: Invalid NumberType=");

	type = RmReadString(rm, L"StringType", L"FILENAME");
	if (_wcsicmp(type, L"FILEPATH") == 0)
		child->strtype = FILEPATH;
	else if (_wcsicmp(type, L"FILENAME") == 0)
		child->strtype = FILENAME;
	else
		RmLog(LOG_ERROR, L"AppVolume.dll: Invalid StringType=");
}

PLUGIN_EXPORT double Update(void* data)
{
	ChildMeasure* child = (ChildMeasure*)data;
	ParentMeasure* parent = child->parent;

	if (!parent)
		return 0.0;

	if (parent->initErr)
	{
		//Try again if calling InitCom() fail in Initialize()
		parent->initErr = !InitCom(parent);
	}

	if (child->isParent)
	{
		if (!UpdateList(parent))
		{
			return 0.0;
		}
	}	
	else
	{
		try
		{
			if (child->indexType == APPNAME)
			{
				for (int i = 0; i < parent->SessionCollection.size(); i++)
				{
					if (_wcsicmp(parent->SessionCollection.at(i)->appName.c_str(), 
						child->indexApp.c_str()) == 0)
					{
						child->indexNum = i;
						break;
					}
				}
			}
			switch (child->numtype)
			{
			case VOLUME:
			{
				float vol = 0.0;
				parent->SessionCollection.at(child->indexNum)->volume->GetMasterVolume(&vol);
				return (parent->SessionCollection.at(child->indexNum)->mute ? -1 : (double)vol);
			}

			case PEAK:
			{
				float peak = 0.0;
				parent->SessionCollection.at(child->indexNum)->peak->GetPeakValue(&peak);
				return (double)peak;
			}
			}
		}
		catch (const std::out_of_range&)
		{
			return 0.0;
		}
	}

	return (double)parent->SessionCollection.size();
}

PLUGIN_EXPORT LPCWSTR GetString(void* data)
{
	ChildMeasure* child = (ChildMeasure*)data;
	ParentMeasure* parent = child->parent;

	if (!parent || parent->initErr)
		return L"Error";

	std::wstring stringValue = L"";

	if (child->isParent)
	{
		stringValue = deviceName;
	}
		
	else
	{
		try
		{
			switch (child->strtype)
			{
			case FILEPATH:
				stringValue = parent->SessionCollection.at(child->indexNum)->appPath;
			case FILENAME:
				stringValue = parent->SessionCollection.at(child->indexNum)->appName;
			}
		}
		catch (const std::out_of_range&)
		{
			return L"Index is out of range";
		}
	}

	return stringValue.c_str();
}

PLUGIN_EXPORT LPCWSTR GetVolumeFromIndex(void* data, const int argc, const WCHAR* argv[])
{
	ChildMeasure* child = (ChildMeasure*)data;
	if (argc == 1)
	{
		int index = _wtoi(argv[0]);
		if (index == 0)
		{
			RmLogF(child->parent->rm, 1, L"AppVolume.dll - GetVolumeFromIndex(%s): Incorrect type or out of range. Please use integer and >= 1.", argv[0]);
			return L"0";
		}
		try
		{
			float vol = 0.0;
			child->parent->SessionCollection.at(index)->volume->GetMasterVolume(&vol);
			vol = (child->parent->SessionCollection.at(index - 1)->mute ? -1 : vol);
			return std::to_wstring(vol).c_str();
		}
		catch (const std::out_of_range&)
		{
			RmLogF(child->parent->rm, 1, L"AppVolume.dll - GetVolumeFromIndex(%d): Index is out of range", index);
			return L"0";
		}
	}
	RmLog(1, L"AppVolume.dll - GetVolumeFromIndex(...): Too many parameters. Please use one integer and >= 1.");
	return L"0";
}

PLUGIN_EXPORT LPCWSTR GetPeakFromIndex(void* data, const int argc, const WCHAR* argv[])
{
	ChildMeasure* child = (ChildMeasure*)data;
	if (argc == 1)
	{
		int index = _wtoi(argv[0]);
		if (index == 0)
		{
			RmLogF(child->parent->rm, 1, L"AppVolume.dll - GetPeakFromIndex(%s): Incorrect type or out of range. Please use integer and >= 1.", argv[0]);
			return L"0";
		}
		try
		{
			float peak = 0.0;
			child->parent->SessionCollection.at(index - 1)->peak->GetPeakValue(&peak);
			return std::to_wstring(peak).c_str();
		}
		catch (const std::out_of_range&)
		{
			RmLogF(child->parent->rm, 1, L"AppVolume.dll - GetPeakFromIndex(%d): Index is out of range", index);
			return L"0";
		}
	}
	RmLog(1, L"AppVolume.dll - GetPeakFromIndex(...): Too many parameters. Please use one integer and >= 1.");
	return L"0";
}

PLUGIN_EXPORT LPCWSTR GetFilePathFromIndex(void* data, const int argc, const WCHAR* argv[])
{
	ChildMeasure* child = (ChildMeasure*)data;
	if (argc == 1)
	{
		int index = _wtoi(argv[0]);
		if (index == 0)
		{
			RmLogF(child->parent->rm, 1, L"AppVolume.dll - GetFilePathFromIndex(%s): Incorrect type or out of range. Please use integer and >= 1.", argv[0]);
			return L"0";
		}
		try
		{
			return child->parent->SessionCollection.at(index - 1)->appPath.c_str();
		}
		catch (const std::out_of_range&)
		{
			RmLogF(child->parent->rm, 1, L"AppVolume.dll - GetFilePathFromIndex(%d): Index is out of range", index);
			return L"0";
		}
	}
	RmLog(1, L"AppVolume.dll - GetFilePathFromIndex(...): Too many parameters. Please use one integer and >= 1.");
	return L"0";
}

PLUGIN_EXPORT LPCWSTR GetFileNameFromIndex(void* data, const int argc, const WCHAR* argv[])
{
	ChildMeasure* child = (ChildMeasure*)data;
	if (argc == 1)
	{
		int index = _wtoi(argv[0]);
		if (index == 0)
		{
			RmLogF(child->parent->rm, 1, L"AppVolume.dll - GetFileNameFromIndex(%s): Incorrect type or out of range. Please use integer and >= 1.", argv[0]);
			return L"0";
		}
		try
		{
			return child->parent->SessionCollection.at(index - 1)->appName.c_str();
		}
		catch (const std::out_of_range&)
		{
			RmLogF(child->parent->rm, 1, L"AppVolume.dll - GetFileNameFromIndex(%d): Index is out of range", index);
			return L"0";
		}
	}
	RmLog(1, L"AppVolume.dll - GetFileNameFromIndex(...): Too many parameters. Please use one integer and >= 1.");
	return L"0";
}

PLUGIN_EXPORT LPCWSTR GetVolumeFromAppName(void* data, const int argc, const WCHAR* argv[])
{
	ChildMeasure* child = (ChildMeasure*)data;
	if (argc == 1)
	{
		for (int i = 0; i < child->parent->SessionCollection.size(); i++)
		{
			if (_wcsicmp(child->parent->SessionCollection.at(i)->appName.c_str(),
				argv[0]) == 0)
			{
				try
				{
					float vol = 0.0;
					child->parent->SessionCollection.at(i)->volume->GetMasterVolume(&vol);
					vol = (child->parent->SessionCollection.at(i)->mute ? -1 : vol);
					return std::to_wstring(vol).c_str();
				}
				catch (const std::out_of_range&)
				{
					RmLogF(child->parent->rm, 1, L"AppVolume.dll - GetVolumeFromAppName(%s): Could not find app.", argv[0]);
					return L"0";
				}
			}
		}
		RmLogF(child->parent->rm, 1, L"AppVolume.dll - GetVolumeFromAppName(%s): Could not find app.", argv[0]);
		return L"0";
	}
	RmLog(1, L"AppVolume.dll - GetVolumeFromAppName(...): Too many parameters. Please use one string.");
	return L"0";
}

PLUGIN_EXPORT LPCWSTR GetPeakFromAppName(void* data, const int argc, const WCHAR* argv[])
{
	ChildMeasure* child = (ChildMeasure*)data;
	if (argc == 1)
	{
		for (int i = 0; i < child->parent->SessionCollection.size(); i++)
		{
			if (_wcsicmp(child->parent->SessionCollection.at(i)->appName.c_str(),
				argv[0]) == 0)
			{
				try
				{
					float peak = 0.0;
					child->parent->SessionCollection.at(i)->peak->GetPeakValue(&peak);
					return std::to_wstring(peak).c_str();
				}
				catch (const std::out_of_range&)
				{
					RmLogF(child->parent->rm, 1, L"AppVolume.dll - GetPeakFromAppName(%s): Could not find app.", argv[0]);
					return L"0";
				}
			}
		}
		RmLogF(child->parent->rm, 1, L"AppVolume.dll - GetPeakFromAppName(%s): Could not find app.", argv[0]);
		return L"0";
	}
	RmLog(1, L"AppVolume.dll - GetPeakFromAppName(...): Too many parameters. Please use one string.");
	return L"0";
}

//Based on Win7AudioPlugin
PLUGIN_EXPORT void ExecuteBang(void* data, LPCWSTR args)
{
	ChildMeasure* child = (ChildMeasure*)data;
	ParentMeasure* parent = child->parent;
	std::wstring wholeBang = args;

	size_t pos = wholeBang.find(' ');
	if (pos != -1)
	{
		std::wstring bang = wholeBang.substr(0, pos);
		wholeBang.erase(0, pos + 1);
		if (_wcsicmp(bang.c_str(), L"SetVolume") == 0)
		{
			float volume = 0;
			if (1 == swscanf_s(wholeBang.c_str(), L"%f", &volume))
			{
				if (parent->SessionCollection.at(child->indexNum)->volume)
				{
					volume = (volume < 0 ? 0 : (volume > 100 ? 100 : volume)) / 100;
					if (!SUCCEEDED(parent->SessionCollection.at(child->indexNum)->volume->SetMasterVolume(volume, NULL)))
						RmLog(LOG_ERROR, L"AppVolume.dll: Error setting volume");
				}
				else
					RmLog(LOG_ERROR, L"AppVolume.dll: Error getting app volume");
			}
			else
				RmLog(LOG_WARNING, L"AppVolume.dll: Incorrect number of arguments for bang");
		}
		else if (_wcsicmp(bang.c_str(), L"ChangeVolume") == 0)
		{
			float offset = 0;
			float volume = 0;
			if (1 == swscanf_s(wholeBang.c_str(), L"%f", &offset))
			{
				if (parent->SessionCollection.at(child->indexNum)->volume)
				{
					parent->SessionCollection.at(child->indexNum)->volume->GetMasterVolume(&volume);
					volume += offset / 100;
					volume = volume < 0 ? 0 : (volume > 1 ? 1 : volume);
					if (!SUCCEEDED(parent->SessionCollection.at(child->indexNum)->volume->SetMasterVolume(volume, NULL)))
						RmLog(LOG_ERROR, L"AppVolume.dll: Error changing volume");
				}
				else
					RmLog(LOG_ERROR, L"AppVolume.dll: Error getting app volume");
			}
			else
				RmLog(LOG_WARNING, L"AppVolume.dll: Incorrect number of arguments for bang");
		}
		else
		{
			RmLog(LOG_WARNING, L"AppVolume.dll: Unknown bang");
		}
	}
	else if (_wcsicmp(wholeBang.c_str(), L"ToggleMute") == 0)
	{
		if (parent->SessionCollection.at(child->indexNum)->volume)
		{
			if (!SUCCEEDED(parent->SessionCollection.at(child->indexNum)->volume->SetMute(!parent->SessionCollection.at(child->indexNum)->mute, NULL)))
				RmLog(LOG_ERROR, L"AppVolume.dll: Error toggling mute app");
		}
		else
			RmLog(LOG_ERROR, L"AppVolume.dll: Error getting app mute status");
	}
	else if (_wcsicmp(wholeBang.c_str(), L"Mute") == 0)
	{
		if (parent->SessionCollection.at(child->indexNum)->volume)
		{
			if (!SUCCEEDED(parent->SessionCollection.at(child->indexNum)->volume->SetMute(true, NULL)))
				RmLog(LOG_ERROR, L"AppVolume.dll: Error muting app");
		}
		else
			RmLog(LOG_ERROR, L"AppVolume.dll: Error getting app mute status");
	}
	else if (_wcsicmp(wholeBang.c_str(), L"Unmute") == 0)
	{
		if (parent->SessionCollection.at(child->indexNum)->volume)
		{
			if (!SUCCEEDED(parent->SessionCollection.at(child->indexNum)->volume->SetMute(false, NULL)))
				RmLog(LOG_ERROR, L"AppVolume.dll: Error unmuting app");
		}
		else
			RmLog(LOG_ERROR, L"AppVolume.dll: Error getting app mute status");
	}
	else if (_wcsicmp(wholeBang.c_str(), L"Update") == 0)
		Reload(data, parent->rm, NULL);
	else
		RmLog(LOG_WARNING, L"AppVolume.dll: Unknown bang");
}
PLUGIN_EXPORT void Finalize(void* data)
{
	ChildMeasure* child = (ChildMeasure*)data;
	ParentMeasure* parent = child->parent;
	if (parent && parent->ownerChild == child)
	{
		std::vector<ParentMeasure*>::iterator iter = std::find(g_ParentMeasures.begin(), g_ParentMeasures.end(), parent);
		g_ParentMeasures.erase(iter);
		parent->SessionCollection.clear();
		delete parent;
	}
	SAFE_RELEASE(pEnumerator);

	deviceName = L"";
	delete child;
	if (com_initialized) CoUninitialize();
	com_initialized = FALSE;
}