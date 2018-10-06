#pragma once
#include <windows.h>

#include <Endpointvolume.h>
#include <Audiopolicy.h>
#include <Mmdeviceapi.h>
#include <Shlwapi.h>
#include <Psapi.h>
#include <Functiondiscoverykeys_devpkey.h>

#include <string>
#include <vector>
#include <Strsafe.h>
#include <algorithm>

#define SAFE_RELEASE(punk)  \
			  if ((punk) != nullptr) { (punk)->Release(); (punk) = nullptr; }

const static CLSID CLSID_MMDeviceEnumerator = __uuidof(MMDeviceEnumerator);
const static IID IID_IMMDeviceEnumerator = __uuidof(IMMDeviceEnumerator);
const static IID IID_SessionManager2 = __uuidof(IAudioSessionManager2);

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

class AppSession
{
public:
    std::vector<IAudioMeterInformation *> peak;
    std::vector<ISimpleAudioVolume *> volume;
    std::vector<DWORD> processID;
    BOOL mute;
    std::wstring appPath = L"";
    std::wstring appName = L"";
    GUID groupID;

    double GetVolume();
    double GetPeak();
    BOOL GetMute();
    LPCWSTR GetPath();
    LPCWSTR GetName();
};

class ParentMeasure
{
public:
    static BOOL isCOMInitialized;
    static IMMDeviceEnumerator * pEnumerator;

    void* skin;
    void* rm;

    BOOL initError;

    LPCWSTR name;
    std::wstring deviceName;

    BOOL ignoreSystemSound;
    std::vector<AppSession> sessionCollection;
    std::vector<std::wstring> excludeList;

    ParentMeasure() :
        skin(nullptr),
        rm(nullptr),
        initError(FALSE),
        name(nullptr),
        ignoreSystemSound(true)
    {};

    ~ParentMeasure();

    BOOL InitializeCOM();
    BOOL UpdateList();
    void ClearSessions();
};

class ChildMeasure
{
public:
    ParentMeasure* parent;
    void* rm;

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
    {};

    BOOL IsOutOfRange();
    int GetIndexFromName(std::wstring name);
    double GetAppVolume();
    double GetAppPeak();
    BOOL GetAppMute();
    LPCWSTR GetAppPath();
    LPCWSTR GetAppName();
};

// Utilities
void SeparateList(LPCWSTR list, std::vector<std::wstring> &vectorList);
BOOL IsValidParent(ChildMeasure * child, LPCWSTR functionName, LPCWSTR functionArg);
BOOL IsValidIndex(int index, ChildMeasure * child, LPCWSTR functionName, LPCWSTR functionArg);
