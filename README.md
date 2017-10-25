# AppVolume Plugin for Rainmeter

A plugin that extend Rainmeter functionality: Get apps volume and peak level, control apps volume and mute.

## Parent measure
Option:  
- `IgnoreSystemSound` *(default = 1)*  
  System sound is Windows notification sound.  
  Set to 0 to include System Sound. Set to 1 to skip System Sound.

Number value: Total number of apps. You can use this value to generate enough child measures and prevent out of range error.  
String value: Current device name.

Example:
```ini
[AppVolumeParent]
Measure = Plugin
Plugin = AppVolume
IgnoreSystemSound = 1
```

## Child measure
Options:
- `Parent`  
  You have to set this to Parent measure name.  
  
- `Index` *(default = 0)*  
  Index of app you want to get infomation and control. It has to be in range from 1 to number value of Parent measure.  
  
- `AppName`  
  Specific name of app you want to get infomation and control. You need to include app name and its extension.  
  Eg: `Spotify.exe`, `AIMP.exe`  
  If this option is not empty and `Index` is also set, it overides `Index` option.  
  
- `NumberType` *(default = volume)*  
  Value you want to return in number value of measure.
  Valid options are `volume` or `peak`.  
  
- `StringType` *(default = filename)*  
  Value you want to return in string value of measure.
  Valid options are `filename` or `filepath`.  
  
Example:  
```ini
[AppIndex2]
Measure = Plugin
Plugin = AppVolume
Parent = AppVolumeParent
Index = 2
NumType = Peak
StringType = Filename
```

```ini
[FoobarVolume]
Measure = Plugin
Plugin = AppVolume
Parent = AppVolumeParent
AppName = Foobar.exe
NumType = Volume
StringType = Filepath
```

## Bangs
Both parent and child:
- `!CommandMeasure MeasureName "Update"`  
  You can use this bang after changing measure option so you do not have to set `DynamicVariables = 1`
  Eg: `LeftMouseUpAction = [!SetOption AppIndex2 NumType Volume][!CommandMeasure AppIndex2 "Update"]`
  
Only child:
- `!CommandMeasure MeasureName "SetVolume x"`
  Set app volume to *x* (from 0 to 100).  
  
- `!CommandMeasure MeasureName "ChangeVolume x"`
  Change the volume by *x*. You can use negative numbers to decrease volume.  

- `!CommandMeasure MeasureName "Mute"`
  Mute app.
  
- `!CommandMeasure MeasureName "UnMute"`
  Unmute app.
 
- `!CommandMeasure MeasureName "ToggleMute"`
  Toggle mute app.
  
## Section variable
Only available in Rainmeter version >= 4.1  
An additional way to get app volume and peak by index or app name. `DynamicVariables = 1` is required in where you use these variables.

`[ParentMeasureName:GetVolumeFromIndex(x)]`  
`[ParentMeasureName:GetPeakFromIndex(x)]`  
`[ParentMeasureName:GetFileNameFromIndex(x)]`  
`[ParentMeasureName:GetFilePathFromIndex(x)]`  
    *x* is from 1 to number value of Parent measure.  

`[ParentMeasureName:GetVolumeFromAppName(name)]`  
`[ParentMeasureName:GetPeakFromAppName(name)]`  
    *name* is name of app you want to get. You need to include app name and its extension.  
    
Example:
```ini
[AppVolumeParent]
Measure = Plugin
Plugin = AppVolume
IgnoreSoundSystem = 0

[Calc_SpotifyVolume]
Measure = Calc
Formula = [AppVolumeParent:GetVolumeFromAppName(spotify.exe)] * 100
DynamicVariables = 1

[Meter_SpotifyVolume]
Meter = String
MeasureName = Calc_SpotifyVolume
Text = Spotify: %1%
FontSize = 40
FontColor = 1FD662
AntiAlias = 1
```

## AppVolume example skin
In release page, I included an example skin pack, you can download, examine and then make your own skin.  
![ExampleSkinDemo](https://raw.githubusercontent.com/khanhas/AppVolumePlugin/master/AppVolumeExampleSkin/demo.png)

## Credit:
Big thanks to [theAzack9](https://github.com/TheAzack9) and [tjhrulz](https://github.com/tjhrulz) who helped me get through some C++ stuff that I'm too noob to understand. I can't finish this one without them.
