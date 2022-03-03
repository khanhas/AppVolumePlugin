#pragma once
#include <windows.h>
#include <commctrl.h>
#include <commoncontrols.h>

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

enum class IconSize
{
    IS_SMALL = 1,	// 16x16
    IS_MEDIUM = 0,	// 32x32
    IS_LARGE = 2,	// 48x48
    IS_EXLARGE = 4	// 256x256
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
    std::wstring saveIcons;
    IconSize iconSize;

    ParentMeasure() :
        skin(nullptr),
        rm(nullptr),
        initError(FALSE),
        name(nullptr),
        ignoreSystemSound(true),
        saveIcons(),
        iconSize(IconSize::IS_LARGE)
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
        parent(),
        rm(),
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
void GetIcon(std::wstring filePath, const std::wstring& iconPath, IconSize iconSize);
bool SaveIcon(HICON hIcon, FILE* fp);