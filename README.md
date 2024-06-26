# FRAM (Fisheries Regulation Assement Model)
## Binary
The latest compiled version of FRAM can be downloaded [here](bin/Debug). Documentation on FRAM can be found on the [documentation page](https://framverse.github.io/fram_doc/index.html).

## Installing for Developement
Requirements:

1. Visual Studio Community [Download](https://visualstudio.microsoft.com/free-developer-offers/).

2. [Github](https://www.github.com) (Optional)

### Visual Studio Installation Process

Visual Studio can be configured to work as an IDE (interactive development environment) for multiple languages. Working in FRAM's source code requires Visual Studio to be configured to work with [Visual Basic / .NET](https://en.wikipedia.org/wiki/Visual_Basic_(.NET)).

When prompted in the installation, `.NET desktop development` at minimum needs to be chosen. 

![Visual Studio Install Screen](img/vs_install_screen.png?raw=true)

The rest of the installation can be completed by repeatedly clicking the `Install` button.


### Downloading FRAM Source Code
Downloading the source code from Github can be downloaded in a zip archive directly from Github.  

![Github FRAM Download](img/gh_dl.png?raw=true "Download FRAM")

The source code can also be download via the normal [Github cloning workflow](https://docs.github.com/en/repositories/creating-and-managing-repositories/cloning-a-repository)

### Visual Studio

Visual Studio is an [IDE](https://en.wikipedia.org/wiki/Integrated_development_environment). It allows you to edit, debug, and compile Visual Basic code - along with other languages. 


After Visual Studio is installed navigate to the unzipped FRAM folder and find `FramVS.vbproj`. This is the project file for the source code and opening it either in File Explorer or Visual Studio will open the source code in the IDE.

#### Upgrading

FRAM is currently coded in Visual Basic 6, which is no longer supported by Microsoft. Luckily the newer versions of Visual Studio will upgrade the source the first time it is opened.

![VB Upgrade](/img/vb_upgrade.png?raw=true)

This prompt currently appears only once, and after clicking `OK` a report is displayed which can be then closed.


#### Layout 

![Visual Studio](/img/screen.png?raw=true)

The dubbuger is near the top of the screen. Pressing the `▶ Start` button will run the source code in debug mode. The highlighted view pane on the right is the collection of "forms" (FVS_NAME.vb) associated with the source code and both house the user interface and code.

#### Important Places
##### Variable Initialization
FRAM makes ubiquitous use of global variables throughout the source. These are defined and initialized in `FarmVars.vb` 

##### Calculations
Most of the code handling FRAM's calculations can be found in the `FramCalcs.vb` module, with the `RunCalcs()` subroutine providing the main processing loop.
