Tangram_Lector_Workstation
=========


## Configuration

The Tangram Lector workstation can be configured by applying an app.config file next to the executable file. The following configurations can be made by adding them in the XML file:

``` XML
<?xml version="1.0"?>
<configuration>
	<appSettings>
		<!-- HERE COMES THE CONFIGURATION -->
	</appSettings>
</configuration>
```

Key| Value | Description | Example
------------ | ------------- | ------------- | ------------
DefaultCulture | Culture code name | The default culture used for translating text messages a.s.o. . For default the language of the operating system is retrieved to determine the language. If your language code is not correct or no translation is available the default language is English. | `<add key="-DefaultCulture" value="en-US" />`
StandardVoice | List of names (separated by comma) of voices that should be taken for tts | The voices are taken by the order given in the list. If the first voice is not available or does not meet the language culture the next is taken. If no voice of the list fits the requirements the standard voice of the system is used. | `<add key="StandardVoice" value="ScanSoft Steffi_Full_22kHz, ScanSoft Steffi_Dri40_16kHz" />`
DefaultLogPriority | tud.mci.tangram.LogPriority (ALWAYS, IMPORTANT, MIDDLE, OFTEN, DEBUG) | defines the minimum priority a log messages must have to be written in the log file. Standard is IMPORTANT. | `<add key="DefaultLogPriority" value="debug" />`
BlockedEXT | List of folder names (separated by comma) of extensions that should not been loaded | The extensions will be ignored in the loading process if the folder name (case sensitive) matches a name in this list | `<add key="BlockedEXT" value="ShowOff" />` 


