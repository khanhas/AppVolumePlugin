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
#include <Strsafe.h>
#include <algorithm>

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
	std::vector<std::wstring> excludeList;
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
	std::vector<IAudioMeterInformation *> peak;
	std::vector<ISimpleAudioVolume *> volume;
	std::vector<DWORD> processID;
	BOOL mute;
	std::wstring appPath = L"";
	std::wstring appName = L"";
	GUID groupID;
};

static BOOL com_initialized = FALSE;
const static CLSID CLSID_MMDeviceEnumerator = __uuidof(MMDeviceEnumerator);
const static IID IID_IMMDeviceEnumerator = __uuidof(IMMDeviceEnumerator);
const static IID IID_SessionManager2 = __uuidof(IAudioSessionManager2);
BOOL InitCom(ParentMeasure* measure);
BOOL UpdateList(ParentMeasure* measure);
void SeparateList(LPCWSTR list, std::vector<std::wstring> &vectorList);
IMMDeviceEnumerator * pEnumerator = nullptr;
std::wstring deviceName;

std::vector<ParentMeasure*> g_ParentMeasures;

BOOL InitCom(ParentMeasure* measure)
{
	if (!com_initialized) com_initialized = SUCCEEDED(CoInitialize(0));
	if (!com_initialized)
	{
		RmLog(LOG_ERROR, L"AppVolume.dll: COM initialization failed");
		return false;
	}

	if (FAILED(CoCreateInstance(CLSID_MMDeviceEnumerator, 0, CLSCTX_ALL, IID_IMMDeviceEnumerator, (void**)&pEnumerator)))
	{
		RmLog(LOG_ERROR, L"AppVolume.dll: COM creation failed");
		return false;
	}
	
	if (FAILED(UpdateList(measure)))
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

void SeparateList(LPCWSTR list, std::vector<std::wstring> &vectorList)
{
	std::wstring tempList = list;
	size_t start = 0;
	size_t end = tempList.find(L";", 0);
	
	while (end != std::wstring::npos)
	{
		start = tempList.find_first_not_of(L" \t\r\n", start);
		std::wstring element = tempList.substr(start, end - start);
		element.erase(element.find_last_not_of(L" \t\r\n") + 1);

		if (!element.empty())
		{
			vectorList.push_back(element);
		}
		start = end + 1;
		end = tempList.find(L";", start);
	}

	if (start < tempList.length())
	{
		start = tempList.find_first_not_of(L" \t\r\n", start);
		std::wstring element = tempList.substr(start, tempList.length() - start);
		element.erase(element.find_last_not_of(L" \t\r\n") + 1);

		if (!element.empty())
		{
			vectorList.push_back(element);
		}		
	}
}
BOOL UpdateList(ParentMeasure* measure)
{
	IMMDevice * pDevice = nullptr;
	if (SUCCEEDED(pEnumerator->GetDefaultAudioEndpoint(eRender, eConsole, &pDevice)))
	{
		IAudioSessionManager2 * sessManager = nullptr;
		if (SUCCEEDED(pDevice->Activate(IID_SessionManager2, CLSCTX_ALL, NULL, (void**)&sessManager)))
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
				
				for (auto &sessIter : measure->SessionCollection)
				{
					for (auto &volIter : sessIter->volume)
						SAFE_RELEASE(volIter);
					sessIter->volume.clear();

					for (auto &peakIter : sessIter->peak)
						SAFE_RELEASE(peakIter);
					sessIter->peak.clear();
					sessIter->processID.clear();
				}
				measure->SessionCollection.clear();
				measure->SessionCollection.shrink_to_fit();

				int sessionCount = 0;
				sessEnum->GetCount(&sessionCount);

				//Go through all sessions, group sessions that have same GUID,
				//create new iter for uniquely found GUID.
				for (int i = 0; i < sessionCount; i++)
				{
					AppSession * newAppSession = new AppSession;

					IAudioSessionControl * sessControl = nullptr;
					if (FAILED(sessEnum->GetSession(i, &sessControl)))
					{
						delete newAppSession;
						continue;
					}

					IAudioSessionControl2 * sessControl2 = nullptr;
					if (FAILED(sessControl->QueryInterface(&sessControl2)))
					{
						SAFE_RELEASE(sessControl);
						delete newAppSession;
						continue;
					}

					HRESULT isSystem = sessControl2->IsSystemSoundsSession();
					if (measure->ignoreSystemSound && isSystem == S_OK)
					{
						SAFE_RELEASE(sessControl);
						SAFE_RELEASE(sessControl2);
						delete newAppSession;
						continue;
					}

					DWORD procID;
					sessControl2->GetProcessId(&procID);

					if (procID == 0)
					{
						newAppSession->appPath = newAppSession->appName = L"System Sound";
					}
					else
					{
						HANDLE checkProc = OpenProcess(PROCESS_ALL_ACCESS, false, procID);
						WCHAR procPath[400];
						procPath[0] = 0;
						if (checkProc != NULL)
						{
							GetModuleFileNameEx(checkProc, NULL, procPath, 400);
						}
						CloseHandle(checkProc);

						//Windows 7 doesn't kill session when process is closed.
						//And it leaves various "ghost" processID that can't be get path.
						//Use that to skip these kind of sessions.
						if (procPath[0] == 0)
						{
							SAFE_RELEASE(sessControl);
							SAFE_RELEASE(sessControl2);
							delete newAppSession;
							continue;
						}
						else
						{
							newAppSession->appPath = procPath;
							newAppSession->appName = PathFindFileName(newAppSession->appPath.c_str());
						}

						BOOL found = FALSE;
						for (auto app : measure->excludeList)
						{
							if (_wcsicmp(newAppSession->appName.c_str(), app.c_str()) == 0)
							{
								found = TRUE;
							}
						}
						if (found)
						{
							SAFE_RELEASE(sessControl);
							SAFE_RELEASE(sessControl2);
							delete newAppSession;
							continue;
						}
					}

					GUID gID;
					sessControl->GetGroupingParam(&gID);
					
					BOOL found = FALSE;
					for (auto &check : measure->SessionCollection)
					{
						if (IsEqualGUID(gID, check->groupID))
						{
							ISimpleAudioVolume * iterVolume = nullptr;
							if (SUCCEEDED(sessControl->QueryInterface(&iterVolume)))
							{
								check->volume.push_back(iterVolume);
								BOOL iterMute;
								iterVolume->GetMute(&iterMute);
								check->mute = check->mute && iterMute;
							}

							IAudioMeterInformation * iterPeak = nullptr;
							if (SUCCEEDED(sessControl->QueryInterface(&iterPeak)))
								check->peak.push_back(iterPeak);

							found = TRUE;
							break;
						}
					}

					if (found)
					{
						SAFE_RELEASE(sessControl);
						SAFE_RELEASE(sessControl2);
						delete newAppSession;
						continue;
					}

					newAppSession->groupID = gID;
					ISimpleAudioVolume * iterVolume = nullptr;
					if (SUCCEEDED(sessControl->QueryInterface(&iterVolume)))
					{
						newAppSession->volume.push_back(iterVolume);
						BOOL iterMute;
						iterVolume->GetMute(&iterMute);
						newAppSession->mute = iterMute;
					}
						
					IAudioMeterInformation * iterPeak = nullptr;
					if (SUCCEEDED(sessControl->QueryInterface(&iterPeak)))
						newAppSession->peak.push_back(iterPeak);

					SAFE_RELEASE(sessControl);
					SAFE_RELEASE(sessControl2);

					measure->SessionCollection.push_back(newAppSession);
				}

				SAFE_RELEASE(sessEnum);

				for (auto &iter : measure->SessionCollection)
				{
					float maxVol = 0.0;
					for (auto &volIter : iter->volume)
					{
						float vol;
						volIter->GetMasterVolume(&vol);

						if (vol > maxVol)
							maxVol = vol;
					}

					for (auto &volIter : iter->volume)
					{
						volIter->SetMasterVolume(maxVol, NULL);
						volIter->SetMute(iter->mute, NULL);
					}

				}
			}

			else
			{
				RmLog(LOG_DEBUG, L"AppVolume.dll: Could not enumerate AudioSessionManager");
				return false;
			}
		}
		else
		{
			SAFE_RELEASE(pDevice);
			RmLog(LOG_DEBUG, L"AppVolume.dll: Could not activate AudioSessionManager");
			return false;
		}
	}
	else
	{
		RmLog(LOG_DEBUG, L"AppVolume.dll: Could not get device");
		return false;
	}
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
		child->parent->initErr = !InitCom(child->parent);
		parent->ignoreSystemSound = RmReadInt(rm, L"IgnoreSystemSound", 1) == 1;
		LPCWSTR excluding = RmReadString(rm, L"ExcludeApp", L"");
		if (*excluding)
		{
			SeparateList(excluding, parent->excludeList);
		}
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
				parent->SessionCollection.at(child->indexNum)->volume[0]->GetMasterVolume(&vol);
				return (parent->SessionCollection.at(child->indexNum)->mute ? -1 : (double)vol);
			}

			case PEAK:
			{
				float peak = 0.0;
				for (auto &peakIter : parent->SessionCollection.at(child->indexNum)->peak)
				{
					float sessPeak;
					peakIter->GetPeakValue(&sessPeak);
					if (sessPeak > peak)
					{
						peak = sessPeak;
					}
				}
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

	if (child->isParent)
	{
		return deviceName.c_str();
	}
		
	else
	{
		try
		{
			switch (child->strtype)
			{
			case FILEPATH:
				return parent->SessionCollection.at(child->indexNum)->appPath.c_str();
			case FILENAME:
				return parent->SessionCollection.at(child->indexNum)->appName.c_str();
			}
		}
		catch (const std::out_of_range&)
		{
			return L"Index is out of range";
		}
	}

	return nullptr;
}

PLUGIN_EXPORT LPCWSTR GetVolumeFromIndex(void* data, const int argc, const WCHAR* argv[])
{
	ChildMeasure* child = (ChildMeasure*)data;
	if (argc == 1)
	{
		int index = _wtoi(argv[0]);
		if (index == 0)
		{
			RmLogF(child->parent->rm, LOG_DEBUG, L"AppVolume.dll - GetVolumeFromIndex(%s): Incorrect type or out of range. Please use one integer and >= 1.", argv[0]);
			return L"0";
		}
		try
		{
			float vol = 0.0;
			child->parent->SessionCollection.at(index - 1)->volume[0]->GetMasterVolume(&vol);
			vol = (child->parent->SessionCollection.at(index - 1)->mute ? -1 : vol);
			static WCHAR result[7];
			StringCchPrintf(result, 7, L"%f", vol);
			return result;
		}
		catch (const std::out_of_range&)
		{
			RmLogF(child->parent->rm, LOG_DEBUG, L"AppVolume.dll - GetVolumeFromIndex(%d): Index is out of range", index);
			return L"0";
		}
	}
	RmLog(LOG_DEBUG, L"AppVolume.dll - GetVolumeFromIndex(...): Incorrect number of parameters. Please use one integer and >= 1.");
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
			RmLogF(child->parent->rm, LOG_DEBUG, L"AppVolume.dll - GetPeakFromIndex(%s): Incorrect type or out of range. Please use one integer and >= 1.", argv[0]);
			return L"0";
		}
		try
		{
			float peak = 0.0;
			for (auto &peakIter : child->parent->SessionCollection.at(index-1)->peak)
			{
				float sessPeak;
				peakIter->GetPeakValue(&sessPeak);
				if (sessPeak > peak)
				{
					peak = sessPeak;
				}
			}
			static WCHAR result[7];
			StringCchPrintf(result, 7, L"%f", peak);
			return result;
		}
		catch (const std::out_of_range&)
		{
			RmLogF(child->parent->rm, LOG_DEBUG, L"AppVolume.dll - GetPeakFromIndex(%d): Index is out of range", index);
			return L"0";
		}
	}
	RmLog(LOG_DEBUG, L"AppVolume.dll - GetPeakFromIndex(...): Incorrect number of parameters. Please use one integer and >= 1.");
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
			RmLogF(child->parent->rm, LOG_DEBUG, L"AppVolume.dll - GetFilePathFromIndex(%s): Incorrect type or out of range. Please use one integer and >= 1.", argv[0]);
			return L"0";
		}
		try
		{
			return child->parent->SessionCollection.at(index - 1)->appPath.c_str();
		}
		catch (const std::out_of_range&)
		{
			RmLogF(child->parent->rm, LOG_DEBUG, L"AppVolume.dll - GetFilePathFromIndex(%d): Index is out of range", index);
			return L"0";
		}
	}
	RmLog(1, L"AppVolume.dll - GetFilePathFromIndex(...): Incorrect number of parameters. Please use one integer and >= 1.");
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
			RmLogF(child->parent->rm, LOG_DEBUG, L"AppVolume.dll - GetFileNameFromIndex(%s): Incorrect type or out of range. Please use one integer and >= 1.", argv[0]);
			return L"0";
		}
		try
		{
			return child->parent->SessionCollection.at(index - 1)->appName.c_str();
		}
		catch (const std::out_of_range&)
		{
			RmLogF(child->parent->rm, LOG_DEBUG, L"AppVolume.dll - GetFileNameFromIndex(%d): Index is out of range", index);
			return L"0";
		}
	}
	RmLog(LOG_DEBUG, L"AppVolume.dll - GetFileNameFromIndex(...): Incorrect number of parameters. Please use one integer and >= 1.");
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
					child->parent->SessionCollection.at(i)->volume[0]->GetMasterVolume(&vol);
					vol = (child->parent->SessionCollection.at(i)->mute ? -1 : vol);
					static WCHAR result[7];
					StringCchPrintf(result, 7, L"%f", vol);
					return result;
				}
				catch (const std::out_of_range&)
				{
					RmLogF(child->parent->rm, LOG_DEBUG, L"AppVolume.dll - GetVolumeFromAppName(%s): Could not find app.", argv[0]);
					return L"0";
				}
			}
		}
		RmLogF(child->parent->rm, LOG_DEBUG, L"AppVolume.dll - GetVolumeFromAppName(%s): Could not find app.", argv[0]);
		return L"0";
	}
	RmLog(LOG_DEBUG, L"AppVolume.dll - GetVolumeFromAppName(...): Incorrect number of parameters. Please use one string.");
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
					for (auto &peakIter : child->parent->SessionCollection.at(i)->peak)
					{
						float sessPeak;
						peakIter->GetPeakValue(&sessPeak);
						if (sessPeak > peak)
						{
							peak = sessPeak;
						}
					}
					static WCHAR result[7];
					StringCchPrintf(result, 7, L"%f", peak);
					return result;
				}
				catch (const std::out_of_range&)
				{
					RmLogF(child->parent->rm, LOG_DEBUG, L"AppVolume.dll - GetPeakFromAppName(%s): Could not find app.", argv[0]);
					return L"0";
				}
			}
		}
		RmLogF(child->parent->rm, LOG_DEBUG, L"AppVolume.dll - GetPeakFromAppName(%s): Could not find app.", argv[0]);
		return L"0";
	}
	RmLog(LOG_DEBUG, L"AppVolume.dll - GetPeakFromAppName(...): Incorrect number of parameters. Please use one string.");
	return L"0";
}

//Based on NowPlaying plugin
PLUGIN_EXPORT void ExecuteBang(void* data, LPCWSTR args)
{
	ChildMeasure* child = (ChildMeasure*)data;
	ParentMeasure* parent = child->parent;

	if (_wcsicmp(args, L"Mute") == 0)
	{
		try
		{
			for (auto &volIter : parent->SessionCollection.at(child->indexNum)->volume)
			{
				if (FAILED(volIter->SetMute(true, NULL)))
					throw "error";
			}
		}
		catch (...)
		{
			RmLog(LOG_ERROR, L"AppVolume.dll: Error muting app");
		}
	}
	else if (_wcsicmp(args, L"UnMute") == 0)
	{
		try
		{
			for (auto &volIter : parent->SessionCollection.at(child->indexNum)->volume)
			{
				if (FAILED(volIter->SetMute(false, NULL)))
					throw "error";
			}
		}
		catch (...)
		{
			RmLog(LOG_ERROR, L"AppVolume.dll: Error unmuting app");
		}
	}
	else if (_wcsicmp(args, L"ToggleMute") == 0)
	{
		try
		{
			BOOL curMute = parent->SessionCollection.at(child->indexNum)->mute;
			for (auto &volIter : parent->SessionCollection.at(child->indexNum)->volume)
			{
				if (FAILED(volIter->SetMute(!curMute, NULL)))
					throw "error";
			}
			parent->SessionCollection.at(child->indexNum)->mute = !curMute;
		}
		catch (...)
		{
			RmLog(LOG_ERROR, L"AppVolume.dll: Error toggling mute app");
		}
	}
	else if (_wcsicmp(args, L"Update") == 0)
	{
		Reload(data, parent->rm, NULL);
	}
	else
	{
		LPCWSTR arg = wcschr(args, L' ');
		if (_wcsnicmp(args, L"SetVolume", 9) == 0)
		{
			try
			{
				float argNum = _wtof(arg);
				float volume = 0.0;
				if (arg[1] == L'+' || arg[1] == L'-')
				{
					// Relative to current volume
					if (FAILED(parent->SessionCollection.at(child->indexNum)->volume[0]->GetMasterVolume(&volume)))
						throw "error";

					volume += argNum / 100;
				}
				else
					volume = argNum / 100;
				if (volume < 0)
					volume = 0.0;
				else if (volume > 1)
					volume = 1.0;
				BOOL curMute;
				parent->SessionCollection.at(child->indexNum)->volume[0]->GetMute(&curMute);
				for (auto &volIter : parent->SessionCollection.at(child->indexNum)->volume)
				{
					if (FAILED(volIter->SetMasterVolume(volume, NULL)))
						throw "error";
					if (curMute)
						if (FAILED(volIter->SetMute(FALSE, NULL)))
							throw "error";
				}
					
			}
			catch (...)
			{
				RmLog(LOG_ERROR, L"AppVolume.dll: Error setting volume");
			}
		}
		else
			RmLog(LOG_WARNING, L"AppVolume.dll: Unknown bang");
	}
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
		SAFE_RELEASE(pEnumerator);
		delete parent;
	}

	delete child;
	if (com_initialized) CoUninitialize();
	com_initialized = FALSE;
}