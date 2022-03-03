#include "AppVolume.h"
#include "API\RainmeterAPI.h"
#include <stdexcept>

#define INVALID_FILE L"/<>\\"

#pragma pack(push, 2)
typedef struct	// 16 bytes
{
    BYTE        bWidth;				// Width, in pixels, of the image
    BYTE        bHeight;			// Height, in pixels, of the image
    BYTE        bColorCount;		// Number of colors in image (0 if >=8bpp)
    BYTE        bReserved;			// Reserved ( must be 0)
    WORD        wPlanes;			// Color Planes
    WORD        wBitCount;			// Bits per pixel
    DWORD       dwBytesInRes;		// How many bytes in this resource?
    DWORD       dwImageOffset;		// Where in the file is this image?
} ICONDIRENTRY, * LPICONDIRENTRY;

typedef struct	// 22 bytes
{
    WORD           idReserved;		// Reserved (must be 0)
    WORD           idType;			// Resource Type (1 for icons)
    WORD           idCount;			// How many images?
    ICONDIRENTRY   idEntries[1];	// An entry for each image (idCount of 'em)
} ICONDIR, * LPICONDIR;
#pragma pack(pop)

BOOL ParentMeasure::isCOMInitialized = FALSE;
IMMDeviceEnumerator * ParentMeasure::pEnumerator = nullptr;

static std::vector<ParentMeasure*> parentCollection;

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

ParentMeasure::~ParentMeasure()
{
    ClearSessions();
}

BOOL ParentMeasure::InitializeCOM()
{
    if (!isCOMInitialized)
        isCOMInitialized = SUCCEEDED(CoInitialize(0));

    if (!isCOMInitialized)
    {
        RmLog(rm, LOG_ERROR, L"AppVolume.dll: COM initialization failed");
        return false;
    }

    if (pEnumerator == nullptr
     && FAILED(CoCreateInstance(
            CLSID_MMDeviceEnumerator,
            0,
            CLSCTX_ALL,
            IID_IMMDeviceEnumerator,
            (void**)&pEnumerator
        )))
    {
        RmLog(rm, LOG_ERROR, L"AppVolume.dll: COM creation failed");
        return false;
    }

    if (FAILED(UpdateList()))
        return false;

    return true;
}

void ParentMeasure::ClearSessions()
{
    for (auto &sessIter : sessionCollection)
    {
        for (auto &volIter : sessIter.volume)
            SAFE_RELEASE(volIter);

        sessIter.volume.clear();

        for (auto &peakIter : sessIter.peak)
            SAFE_RELEASE(peakIter);

        sessIter.peak.clear();
        sessIter.processID.clear();
    }

    sessionCollection.clear();
    sessionCollection.shrink_to_fit();
}

BOOL ParentMeasure::UpdateList()
{
    IMMDevice * pDevice = nullptr;
    if (FAILED(pEnumerator->GetDefaultAudioEndpoint(eRender, eConsole, &pDevice)))
    {
        RmLog(rm, LOG_DEBUG, L"AppVolume.dll: Could not get device");
        return false;
    }

    IAudioSessionManager2 * sessManager = nullptr;
    if (FAILED(pDevice->Activate(IID_SessionManager2, CLSCTX_ALL, NULL, (void**)&sessManager)))
    {
        SAFE_RELEASE(pDevice);
        RmLog(rm, LOG_DEBUG, L"AppVolume.dll: Could not activate AudioSessionManager");
        return false;
    }

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

    if (FAILED(sessManager->GetSessionEnumerator(&sessEnum)))
    {
        RmLog(rm, LOG_DEBUG, L"AppVolume.dll: Could not enumerate AudioSessionManager");
        SAFE_RELEASE(sessManager);
        return false;
    }

    SAFE_RELEASE(sessManager);

    ClearSessions();

    int sessionCount = 0;
    sessEnum->GetCount(&sessionCount);

    //Go through all sessions, group sessions that have same GUID,
    //create new iter for uniquely found GUID.
    for (int i = 0; i < sessionCount; i++)
    {
        IAudioSessionControl * sessControl = nullptr;
        if (FAILED(sessEnum->GetSession(i, &sessControl)))
        {
            continue;
        }

        IAudioSessionControl2 * sessControl2 = nullptr;
        if (FAILED(sessControl->QueryInterface(&sessControl2)))
        {
            SAFE_RELEASE(sessControl);
            continue;
        }

        SAFE_RELEASE(sessControl);

        HRESULT isSystem = sessControl2->IsSystemSoundsSession();
        if (ignoreSystemSound && isSystem == S_OK)
        {
            SAFE_RELEASE(sessControl2);
            continue;
        }

        DWORD procID;
        sessControl2->GetProcessId(&procID);

        AppSession newAppSession;

        if (procID == 0)
        {
            newAppSession.appPath = newAppSession.appName = L"System Sound";
        }
        else
        {
            HANDLE checkProc = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, false, procID);
            WCHAR procPath[400];
            DWORD copied = (DWORD)0;
            if (checkProc != NULL)
            {
                copied = GetModuleFileNameEx(checkProc, NULL, procPath, 400);
            }
            CloseHandle(checkProc);

            //Windows 7 doesn't kill session when process is closed.
            //And it leaves various "ghost" processID that can't be get path.
            //Use that to skip these kind of sessions.
            if (copied == (DWORD)0)
            {
                SAFE_RELEASE(sessControl2);
                continue;
            }

            newAppSession.appPath = procPath;
            newAppSession.appName = PathFindFileName(newAppSession.appPath.c_str());

            BOOL found = FALSE;
            for (auto app : excludeList)
            {
                if (_wcsicmp(newAppSession.appName.c_str(), app.c_str()) == 0)
                {
                    found = TRUE;
                }
            }
            if (found)
            {
                SAFE_RELEASE(sessControl2);
                continue;
            }
        }

        GUID gID;
        sessControl2->GetGroupingParam(&gID);

        BOOL found = FALSE;
        for (auto &session : sessionCollection)
        {
            if (!IsEqualGUID(gID, session.groupID))
            {
                continue;
            }

            ISimpleAudioVolume * iterVolume = nullptr;
            if (SUCCEEDED(sessControl2->QueryInterface(&iterVolume)))
            {
                session.volume.push_back(iterVolume);
                BOOL iterMute;
                iterVolume->GetMute(&iterMute);
                session.mute = session.mute && iterMute;
            }

            IAudioMeterInformation * iterPeak = nullptr;
            if (SUCCEEDED(sessControl2->QueryInterface(&iterPeak)))
                session.peak.push_back(iterPeak);

            found = TRUE;
            break;
        }

        if (found)
        {
            SAFE_RELEASE(sessControl2);
            continue;
        }

        newAppSession.groupID = gID;
        ISimpleAudioVolume * iterVolume = nullptr;
        if (SUCCEEDED(sessControl2->QueryInterface(&iterVolume)))
        {
            newAppSession.volume.push_back(iterVolume);
            BOOL iterMute;
            iterVolume->GetMute(&iterMute);
            newAppSession.mute = iterMute;
        }

        IAudioMeterInformation * iterPeak = nullptr;
        if (SUCCEEDED(sessControl2->QueryInterface(&iterPeak)))
            newAppSession.peak.push_back(iterPeak);

        SAFE_RELEASE(sessControl2);

        sessionCollection.push_back(newAppSession);
    }

    SAFE_RELEASE(sessEnum);

    for (auto &session : sessionCollection)
    {
        if (!saveIcons.empty() && !PathFileExists((saveIcons + session.appName + L".png").c_str())) {
            GetIcon(session.appPath, saveIcons + session.appName + L".png", iconSize);
        }
        float maxVol = 0.0;
        for (auto &volIter : session.volume)
        {
            float vol;
            volIter->GetMasterVolume(&vol);

            if (vol > maxVol)
                maxVol = vol;
        }

        for (auto &volIter : session.volume)
        {
            volIter->SetMasterVolume(maxVol, NULL);
            volIter->SetMute(session.mute, NULL);
        }
    }

    return true;
}

BOOL ChildMeasure::IsOutOfRange()
{
    if (indexNum < 0 || indexNum >= (int)parent->sessionCollection.size())
    {
        return TRUE;
    }

    return FALSE;
}

int ChildMeasure::GetIndexFromName(std::wstring name)
{
    for (auto &session : parent->sessionCollection)
    {
        if (_wcsicmp(session.GetName(), name.c_str()) == 0)
        {
            return (int)(&session - &parent->sessionCollection[0]);
            break;
        }
    }

    return -1;
}

double ChildMeasure::GetAppVolume()
{
    if (IsOutOfRange())
    {
        return 0.0;
    }

    return parent->sessionCollection.at(indexNum).GetVolume();
}

double ChildMeasure::GetAppPeak()
{
    if (IsOutOfRange())
    {
        return 0.0;
    }

    return parent->sessionCollection.at(indexNum).GetPeak();
}

BOOL ChildMeasure::GetAppMute()
{
    if (IsOutOfRange())
    {
        return FALSE;
    }

    return parent->sessionCollection.at(indexNum).GetMute();
}

LPCWSTR ChildMeasure::GetAppPath()
{
    if (IsOutOfRange())
    {
        return L"";
    }

    return parent->sessionCollection.at(indexNum).GetPath();
}

LPCWSTR ChildMeasure::GetAppName()
{
    if (IsOutOfRange())
    {
        return L"";
    }

    return parent->sessionCollection.at(indexNum).GetName();
}

double AppSession::GetVolume()
{
    float value = 0.0;
    for (auto &volumeIter : volume)
    {
        float sessVol;
        volumeIter->GetMasterVolume(&sessVol);
        if (sessVol > value)
        {
            value = sessVol;
        }
    }

    return (double)value;
}

double AppSession::GetPeak()
{
    float value = 0.0;
    for (auto &peakIter : peak)
    {
        float sessPeak;
        peakIter->GetPeakValue(&sessPeak);
        if (sessPeak > value)
        {
            value = sessPeak;
        }
    }

    return (double)value;
}

BOOL AppSession::GetMute()
{
    return mute;
}

LPCWSTR AppSession::GetPath()
{
    return appPath.c_str();
}

LPCWSTR AppSession::GetName()
{
    return appName.c_str();
}

PLUGIN_EXPORT void Initialize(void** data, void* rm)
{
    ChildMeasure* child = new ChildMeasure;
    *data = child;

    child->rm = rm;

    void* skin = RmGetSkin(rm);

    std::wstring parentName = RmReadString(rm, L"Parent", L"");
    if (parentName.empty())
    {
        child->parent = new ParentMeasure;
        child->parent->name = RmGetMeasureName(rm);
        child->parent->skin = skin;
        parentCollection.push_back(child->parent);
        child->isParent = true;
    }
    else
    {
        // Find parent using name AND the skin handle to be sure that it's the right one
        for (auto &parent : parentCollection)
        {
            if (_wcsicmp(parent->name, parentName.c_str()) == 0 &&
                parent->skin == skin)
            {
                child->parent = parent;
                return;
            }
        }

        RmLog(LOG_ERROR, L"AppVolume.dll: Invalid Parent=");
    }
}

PLUGIN_EXPORT void Reload(void* data, void* rm, double* maxValue)
{
    ChildMeasure* child = (ChildMeasure*)data;
    ParentMeasure* parent = child->parent;
    if (!parent)
        return;

    if (child->isParent)
    {
        parent->InitializeCOM();
        parent->ignoreSystemSound = RmReadInt(rm, L"IgnoreSystemSound", 1) == 1;
        std::wstring saveIcons = RmReadPath(rm, L"IconsPath", L"");
        if (saveIcons.find_last_of(L'\\') != saveIcons.length()) {
            saveIcons += L'\\';
        }
        parent->saveIcons = saveIcons;
        int iconSize = RmReadInt(rm, L"IconSize", 48);
        switch (iconSize)
        {
        case 16:
            parent->iconSize = IconSize::IS_SMALL;
            break;
        case 32:
            parent->iconSize = IconSize::IS_MEDIUM;
            break;
        case 256:
            parent->iconSize = IconSize::IS_EXLARGE;
            break;
        default:
            break;
        }
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

    if (parent->initError)
    {
        //Try again if calling InitializeCOM fails in Initialize
        parent->initError = !parent->InitializeCOM();
    }

    if (child->isParent)
    {
        if (!parent->UpdateList())
        {
            return 0.0;
        }
    }
    else
    {
        try
        {
            if (child->indexType == APPNAME)
                child->indexNum = child->GetIndexFromName(child->indexApp);

            switch (child->numtype)
            {
            case VOLUME:
            {
                if (child->GetAppMute())
                    return -1;

                return child->GetAppVolume();
            }

            case PEAK:
            {
                return child->GetAppPeak();
            }
            }
        }
        catch (const std::out_of_range&)
        {
            return 0.0;
        }
    }

    // Parent number value
    return (double)parent->sessionCollection.size();
}

PLUGIN_EXPORT LPCWSTR GetString(void* data)
{
    ChildMeasure* child = (ChildMeasure*)data;
    ParentMeasure* parent = child->parent;

    if (!parent || parent->initError)
        return L"Error";

    if (child->isParent)
    {
        return parent->deviceName.c_str();
    }
    else
    {
        switch (child->strtype)
        {
        case FILEPATH:
            return child->GetAppPath();
        case FILENAME:
            return child->GetAppName();
        }
    }

    return nullptr;
}

BOOL IsValidParent(ChildMeasure * child, LPCWSTR functionName, LPCWSTR functionArg)
{
    if (!child->parent)
    {
        RmLogF(
            child->rm,
            LOG_DEBUG,
            L"AppVolume.dll - %s(%s): Invalid measure.",
            functionName,
            functionArg);

        return FALSE;
    }

    if (child->parent->initError)
    {
        RmLogF(
            child->rm,
            LOG_DEBUG,
            L"AppVolume.dll - %s(%s): Parent is not initialized.",
            functionName,
            functionArg);

        return FALSE;
    }

    return TRUE;
}

BOOL IsValidIndex(int index, ChildMeasure * child, LPCWSTR functionName, LPCWSTR functionArg)
{
    if (index == 0)
    {
        RmLogF(
            child->rm,
            LOG_DEBUG,
            L"AppVolume.dll - GetVolumeFromIndex(%s): Incorrect type or out of range. Please use one integer and >= 1.",
            functionArg);

        return FALSE;
    }

    if (index < 1 || index >(int)child->parent->sessionCollection.size())
    {
        RmLogF(
            child->rm,
            LOG_DEBUG,
            L"AppVolume.dll - GetVolumeFromIndex(%d): Index is out of range",
            index);

        return FALSE;
    }

    return TRUE;
}

PLUGIN_EXPORT LPCWSTR GetVolumeFromIndex(void* data, const int argc, const WCHAR* argv[])
{
    ChildMeasure* child = (ChildMeasure*)data;
    ParentMeasure* parent = child->parent;

    if (!IsValidParent(child, L"GetVolumeFromIndex", argv[0]))
    {
        return nullptr;
    }

    if (argc == 1)
    {
        int index = _wtoi(argv[0]);

        if (!IsValidIndex(index, child, L"GetVolumeFromIndex", argv[0]))
        {
            return L"0";
        }

        index--;

        AppSession* session = &parent->sessionCollection.at(index);
        double vol = 0.0;
        if (session->GetMute())
            vol = -1;
        else
            vol = session->GetVolume();

        static WCHAR result[7];
        StringCchPrintf(result, 7, L"%f", vol);
        return result;
    }

    RmLogF(child->rm, LOG_DEBUG, L"AppVolume.dll - GetVolumeFromIndex(...): Incorrect number of parameters. Please use one integer and >= 1.");
    return L"0";
}

PLUGIN_EXPORT LPCWSTR GetPeakFromIndex(void* data, const int argc, const WCHAR* argv[])
{
    ChildMeasure* child = (ChildMeasure*)data;
    ParentMeasure* parent = child->parent;

    if (!IsValidParent(child, L"GetPeakFromIndex", argv[0]))
    {
        return nullptr;
    }

    if (argc == 1)
    {
        int index = _wtoi(argv[0]);

        if (!IsValidIndex(index, child, L"GetPeakFromIndex", argv[0]))
        {
            return L"0";
        }

        index--;

        double peak = parent->sessionCollection.at(index).GetPeak();
        static WCHAR result[7];
        StringCchPrintf(result, 7, L"%f", peak);
        return result;
    }
    RmLog(child->rm, LOG_DEBUG, L"AppVolume.dll - GetPeakFromIndex(...): Incorrect number of parameters. Please use one integer and >= 1.");
    return L"0";
}

PLUGIN_EXPORT LPCWSTR GetFilePathFromIndex(void* data, const int argc, const WCHAR* argv[])
{
    ChildMeasure* child = (ChildMeasure*)data;
    ParentMeasure* parent = child->parent;

    if (!IsValidParent(child, L"GetFilePathFromIndex", argv[0]))
    {
        return nullptr;
    }

    if (argc == 1)
    {
        int index = _wtoi(argv[0]);

        if (!IsValidIndex(index, child, L"GetFilePathFromIndex", argv[0]))
        {
            return L"0";
        }

        index--;
        return child->parent->sessionCollection.at(index).GetPath();
    }

    RmLog(child->rm, LOG_DEBUG, L"AppVolume.dll - GetFilePathFromIndex(...): Incorrect number of parameters. Please use one integer and >= 1.");
    return L"0";
}

PLUGIN_EXPORT LPCWSTR GetFileNameFromIndex(void* data, const int argc, const WCHAR* argv[])
{
    ChildMeasure* child = (ChildMeasure*)data;
    ParentMeasure* parent = child->parent;

    if (!IsValidParent(child, L"GetFileNameFromIndex", argv[0]))
    {
        return nullptr;
    }

    if (argc == 1)
    {
        int index = _wtoi(argv[0]);

        if (!IsValidIndex(index, child, L"GetFileNameFromIndex", argv[0]))
        {
            return L"0";
        }

        index--;
        return child->parent->sessionCollection.at(index).GetName();
    }

    RmLog(child->rm, LOG_DEBUG, L"AppVolume.dll - GetFileNameFromIndex(...): Incorrect number of parameters. Please use one integer and >= 1.");
    return L"0";
}

PLUGIN_EXPORT LPCWSTR GetVolumeFromAppName(void* data, const int argc, const WCHAR* argv[])
{
    ChildMeasure* child = (ChildMeasure*)data;
    ParentMeasure* parent = child->parent;

    if (!IsValidParent(child, L"GetVolumeFromAppName", argv[0]))
    {
        return nullptr;
    }

    if (argc == 1)
    {
        int index = child->GetIndexFromName(argv[0]);
        if (index == -1)
        {
            RmLogF(child->parent->rm, LOG_DEBUG, L"AppVolume.dll - GetVolumeFromAppName(%s): Could not find app.", argv[0]);
            return L"0";
        }

        AppSession* session = &parent->sessionCollection.at(index);
        double vol = 0.0;
        if (session->GetMute())
            vol = -1;
        else
            vol = session->GetVolume();

        static WCHAR result[7];
        StringCchPrintf(result, 7, L"%f", vol);
        return result;
    }

    RmLog(child->rm, LOG_DEBUG, L"AppVolume.dll - GetVolumeFromAppName(...): Incorrect number of parameters. Please use one string.");
    return L"0";
}

PLUGIN_EXPORT LPCWSTR GetPeakFromAppName(void* data, const int argc, const WCHAR* argv[])
{
    ChildMeasure* child = (ChildMeasure*)data;
    ParentMeasure* parent = child->parent;

    if (!IsValidParent(child, L"GetPeakFromAppName", argv[0]))
    {
        return nullptr;
    }

    if (argc == 1)
    {
        int index = child->GetIndexFromName(argv[0]);
        if (index == -1)
        {
            RmLogF(child->rm, LOG_DEBUG, L"AppVolume.dll - GetPeakFromAppName(%s): Could not find app.", argv[0]);
            return L"0";
        }

        AppSession* session = &parent->sessionCollection.at(index);
        double peak = session->GetPeak();

        static WCHAR result[7];
        StringCchPrintf(result, 7, L"%f", peak);
        return result;
    }

    RmLogF(child->rm, LOG_DEBUG, L"AppVolume.dll - GetPeakFromAppName(...): Incorrect number of parameters. Please use one string.");
    return L"0";
}

//Based on NowPlaying plugin
PLUGIN_EXPORT void ExecuteBang(void* data, LPCWSTR args)
{
    ChildMeasure* child = (ChildMeasure*)data;
    ParentMeasure* parent = child->parent;
    if (!parent)
    {
        RmLogF(child->rm, LOG_ERROR, L"AppVolume.dll - Bang \"%s\": Invalid measure.", args);
        return;
    }

    if (parent->initError)
    {
        RmLogF(child->rm, LOG_ERROR, L"AppVolume.dll - Bang \"%s\": Parent is not initialized.", args);
        return;
    }

    if (child->isParent)
    {
        if (_wcsicmp(args, L"Update") == 0)
        {
            Reload(data, parent->rm, NULL);
        }
        else
        {
            RmLog(child->rm, LOG_WARNING, L"AppVolume.dll: Unknown bang");
        }
        return;
    }

    if (child->IsOutOfRange())
    {
        RmLog(child->rm, LOG_ERROR, L"AppVolume.dll: Child measure index is out of range");
        return;
    }

    AppSession* session = &parent->sessionCollection.at(child->indexNum);

    if (_wcsicmp(args, L"Mute") == 0)
    {
        try
        {
            for (auto &volIter : session->volume)
            {
                if (FAILED(volIter->SetMute(TRUE, NULL)))
                    throw "error";
            }
        }
        catch (...)
        {
            RmLog(child->rm, LOG_ERROR, L"AppVolume.dll: Error muting app");
        }
    }
    else if (_wcsicmp(args, L"UnMute") == 0)
    {
        try
        {
            for (auto &volIter : session->volume)
            {
                if (FAILED(volIter->SetMute(FALSE, NULL)))
                    throw "error";
            }
        }
        catch (...)
        {
            RmLog(child->rm, LOG_ERROR, L"AppVolume.dll: Error unmuting app");
        }
    }
    else if (_wcsicmp(args, L"ToggleMute") == 0)
    {
        try
        {
            BOOL newMuteState = !session->GetMute();
            for (auto &volIter : session->volume)
            {
                if (FAILED(volIter->SetMute(newMuteState, NULL)))
                    throw "error";
            }

            session->mute = newMuteState;
        }
        catch (...)
        {
            RmLog(child->rm, LOG_ERROR, L"AppVolume.dll: Error toggling mute app");
        }
    }
    else if (_wcsicmp(args, L"Update") == 0)
    {
        Reload(data, child->rm, NULL);
    }
    else
    {
        LPCWSTR arg = wcschr(args, L' ');
        if (_wcsnicmp(args, L"SetVolume", 9) == 0
         || _wcsnicmp(args, L"ChangeVolume", 12) == 0)
        {
            try
            {
                double argNum = _wtof(arg);
                double volume = 0.0;
                if (arg[1] == L'+' || arg[1] == L'-')
                {
                    volume = session->GetVolume();

                    volume += argNum / 100;
                }
                else
                {
                    volume = argNum / 100;
                }

                if (volume < 0)
                    volume = 0.0;
                else if (volume > 1)
                    volume = 1.0;

                BOOL curMute = session->GetMute();

                for (auto &volIter : session->volume)
                {
                    if (FAILED(volIter->SetMasterVolume((float)volume, NULL)))
                        throw "error";

                    if (curMute)
                        if (FAILED(volIter->SetMute(FALSE, NULL)))
                            throw "error";
                }

            }
            catch (...)
            {
                RmLog(child->rm, LOG_ERROR, L"AppVolume.dll: Error setting volume");
            }
        }
        else
        {
            RmLog(child->rm, LOG_WARNING, L"AppVolume.dll: Unknown bang");
        }
    }
}

PLUGIN_EXPORT void Finalize(void* data)
{
    ChildMeasure* child = (ChildMeasure*)data;
    ParentMeasure* parent = child->parent;
    if (parent && child->isParent)
    {
        auto parentIter = std::find(
            parentCollection.begin(),
            parentCollection.end(),
            parent);

        parentCollection.erase(parentIter);

        delete parent;

        if (parentCollection.size() == 0)
        {
            if (ParentMeasure::isCOMInitialized)
            {
                CoUninitialize();
                ParentMeasure::isCOMInitialized = FALSE;
            }

            SAFE_RELEASE(ParentMeasure::pEnumerator);
        }
    }

    delete child;
}


// copied from file view plugin
void GetIcon(std::wstring filePath, const std::wstring& iconPath, IconSize iconSize)
{
    SHFILEINFO shFileInfo;
    HICON icon = nullptr;
    HIMAGELIST* hImageList = nullptr;
    FILE* fp = nullptr;

    // Special case for .url files
    if (filePath.size() > 3 && _wcsicmp(filePath.substr(filePath.size() - 4).c_str(), L".URL") == 0)
    {
        WCHAR buffer[MAX_PATH] = L"";
        GetPrivateProfileString(L"InternetShortcut", L"IconFile", L"", buffer, sizeof(buffer), filePath.c_str());
        if (*buffer)
        {
            std::wstring file = buffer;
            int iconIndex = 0;

            GetPrivateProfileString(L"InternetShortcut", L"IconIndex", L"-1", buffer, sizeof(buffer), filePath.c_str());
            if (buffer != L"-1")
            {
                iconIndex = _wtoi(buffer);
            }

            int size = 48;
            switch (iconSize) {
            case IconSize::IS_SMALL:
                size = 16;
            case IconSize::IS_MEDIUM:
                size = 32;
            case IconSize::IS_EXLARGE:
                size = 256;
            }

            PrivateExtractIcons(file.c_str(), iconIndex, size, size, &icon, nullptr, 1, LR_LOADTRANSPARENT);
        }
    }

    if (icon == nullptr)
    {
        SHGetFileInfo(filePath.c_str(), 0, &shFileInfo, sizeof(shFileInfo), SHGFI_SYSICONINDEX);
        SHGetImageList((int)iconSize, IID_IImageList, (void**)&hImageList);
        ((IImageList*)hImageList)->GetIcon(shFileInfo.iIcon, ILD_TRANSPARENT, &icon);
    }

    errno_t error = _wfopen_s(&fp, iconPath.c_str(), L"wb");
    if (filePath == INVALID_FILE || icon == nullptr || (error == 0 && !SaveIcon(icon, fp)))
    {
        fwrite(iconPath.c_str(), 1, 1, fp);		// Clears previous icon
        fclose(fp);
    }

    DestroyIcon(icon);
}

bool SaveIcon(HICON hIcon, FILE* fp)
{
    ICONINFO iconInfo;
    BITMAP bmColor;
    BITMAP bmMask;
    if (!fp || nullptr == hIcon || !GetIconInfo(hIcon, &iconInfo) ||
        !GetObject(iconInfo.hbmColor, sizeof(bmColor), &bmColor) ||
        !GetObject(iconInfo.hbmMask, sizeof(bmMask), &bmMask))
        return false;

    // support only 16/32 bit icon now
    if (bmColor.bmBitsPixel != 16 && bmColor.bmBitsPixel != 32)
        return false;

    HDC dc = GetDC(nullptr);
    BYTE bmiBytes[sizeof(BITMAPINFOHEADER) + 256 * sizeof(RGBQUAD)];
    BITMAPINFO* bmi = (BITMAPINFO*)bmiBytes;

    // color bits
    memset(bmi, 0, sizeof(BITMAPINFO));
    bmi->bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    GetDIBits(dc, iconInfo.hbmColor, 0, bmColor.bmHeight, nullptr, bmi, DIB_RGB_COLORS);
    int colorBytesCount = bmi->bmiHeader.biSizeImage;
    BYTE* colorBits = new BYTE[colorBytesCount];
    GetDIBits(dc, iconInfo.hbmColor, 0, bmColor.bmHeight, colorBits, bmi, DIB_RGB_COLORS);

    // mask bits
    memset(bmi, 0, sizeof(BITMAPINFO));
    bmi->bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    GetDIBits(dc, iconInfo.hbmMask, 0, bmMask.bmHeight, nullptr, bmi, DIB_RGB_COLORS);
    int maskBytesCount = bmi->bmiHeader.biSizeImage;
    BYTE* maskBits = new BYTE[maskBytesCount];
    GetDIBits(dc, iconInfo.hbmMask, 0, bmMask.bmHeight, maskBits, bmi, DIB_RGB_COLORS);

    ReleaseDC(nullptr, dc);

    // icon data
    BITMAPINFOHEADER bmihIcon;
    memset(&bmihIcon, 0, sizeof(bmihIcon));
    bmihIcon.biSize = sizeof(BITMAPINFOHEADER);
    bmihIcon.biWidth = bmColor.bmWidth;
    bmihIcon.biHeight = bmColor.bmHeight * 2;	// icXOR + icAND
    bmihIcon.biPlanes = bmColor.bmPlanes;
    bmihIcon.biBitCount = bmColor.bmBitsPixel;
    bmihIcon.biSizeImage = colorBytesCount + maskBytesCount;

    // icon header
    ICONDIR dir;
    dir.idReserved = 0;		// must be 0
    dir.idType = 1;			// 1 for icons
    dir.idCount = 1;
    dir.idEntries[0].bWidth = (BYTE)bmColor.bmWidth;
    dir.idEntries[0].bHeight = (BYTE)bmColor.bmHeight;
    dir.idEntries[0].bColorCount = 0;		// 0 if >= 8bpp
    dir.idEntries[0].bReserved = 0;		// must be 0
    dir.idEntries[0].wPlanes = bmColor.bmPlanes;
    dir.idEntries[0].wBitCount = bmColor.bmBitsPixel;
    dir.idEntries[0].dwBytesInRes = sizeof(bmihIcon) + bmihIcon.biSizeImage;
    dir.idEntries[0].dwImageOffset = sizeof(ICONDIR);

    fwrite(&dir, sizeof(dir), 1, fp);
    fwrite(&bmihIcon, sizeof(bmihIcon), 1, fp);
    fwrite(colorBits, colorBytesCount, 1, fp);
    fwrite(maskBits, maskBytesCount, 1, fp);

    // Clean up
    DeleteObject(iconInfo.hbmColor);
    DeleteObject(iconInfo.hbmMask);
    delete[] colorBits;
    delete[] maskBits;

    fclose(fp);

    return true;
}