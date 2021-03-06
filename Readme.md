Stata Editor for Sublime Text 3
===============================
Tutorial available at: [http://sergiocorreia.github.io/StataEditor/](http://sergiocorreia.github.io/StataEditor/)

Original work by [Mattias Nordin](http://sites.google.com/site/econnordin/)

Fork by [Sergio Correia](http://scorreia.com/)

This is an experimental fork for personal use. It is stable but tread with caution.

* Version 0.5.0
* Date: November 5, 2014

Features
--------
This package provides the ability to write and run Stata code from Sublime Text 3 (ST3). Features in this package include:

* Clear syntax highlighting
* Extends Sublime Text "build" feature to support the Stata _do_ (ctrl+b) and _run_ commands (ctrl+shift+b).
* Run only selected code with support for multiple selections
* Two dozen snippets and completions
* When the required metadata is set (see _autocomplete_ snippet), allows variable and dataset autocompletion
* Autocomplete indicates sorting variable
* Variable autocompletion inspects i) datasets in the metadata paths and ii) generate/egen lines, and is available with ctrl+shift+space
* Also it selects variables of the relevant dataset when doing use/merge/etc. This takes advantage of the order of the snippets
* Supports goto-symbol (ctrl+r) for program and block headers (snippet "header")
* Dataset autocompletion inspects i) metadata paths and ii) save lines, and is available after _use_ or _using_ keywords, either automatically or with ctrl+space
* For speed reasons (or if the datasets are not locally available), the autocomplete contents are saved in a JSON file (usually stata-autocomplete.json)
* Access Stata help files by selecting the command for which you want access to the documentation and press F1 (open help file in Sublime Text) or shft+F1 (open help file in Stata). For the latter option, an internet connection is required.
* Automatic expansion of quotes and local variable delimiters
* Shorthand for creation of locals by pressing ctrl+shift+l
* Load a new Stata dataset by selecting a path and pressing ctrl + shift + u (Equivalent to the command "use 'path', clear"). Please, note that your current work will then be lost, so remember to save your dataset!
* Plus all other features that come with ST3!

Misc:

* Added menu and tutorial link

Requirements and Setup
----------------------
This package only works on Windows machines. To use Stata with Sublime Text on OS X, try [Stata Enhanced](https://sublime.wbond.net/packages/Stata%20Enhanced). StataEditor has been tested on Sublime Text 3 together with Stata 13-MP on Windows 7 and Windows 8. I have very briefly tested it on Stata 11 and Stata 12 and it seems to be working, but more testing is needed.

To install the package follow the steps outlined below. You can install StataEditor without Package Control, but in that case you probably already know what to do.

1. Download and install [ST3](http://www.sublimetext.com/3) if you do not already have it installed.

2. Install Package Control. To get Package Control, click [here](https://sublime.wbond.net/installation) and follow the instructions for ST3.

3. Open ST3 and click Preferences -> Package Control. Choose "Install Package" and choose StataEditor from the list.

4. Repeat this step and install the Pywin32 package.

5. If the path to your Stata installation is "C:/Program Files (x86)/Stata13/StataMP-64.exe" you can skip this step. If not, select Preferences -> Package Settings -> StataEditor -> Settings - Default. Copy the content and then go to Preferences -> Package Settings -> StataEditor -> Settings - User and paste your copied text in the new file. Then change the path to where your Stata installation is located (note that you need to use forward slash, "/", instead of backward slash, "\") and save the file. Do not change the content of the Settings - Default file. While this will work temporarily, with the next update your changes will disappear. The content of the Settings - User file will not be overwritten when the package is updated.

6. Finally, to use Stata interactively from ST3, you also need to register the Stata Automation type library. Instructions can be found [here](http://www.stata.com/automation/#createmsapp). Note that I have had to use the Windows Vista instructions for Windows 7 to get Stata Automation to work. Once the Stata Automation type library has been registered, you are good to go!

Known issues
------------
The development of this package is still in beta and may contain bugs, so use at your own risk and make sure you backup your data. When running code from ST3, a new instance of Stata is opened. If you close ST3, then that instance will also close, though it may take around five minutes before that happens. Note that Stata will not ask you whether you want to save the data but will close without warning. **Therefore, do not use an instance of Stata that has been launched from ST3 after ST3 has been closed, as you would risk loosing your unsaved work.** This is true even if you re-launch ST3 as the connection to the old instance of Stata will have been permanently broken. Please let me know if you detect any other bugs or if you have requests for additional featuers. You can contact me at mnordin [at] gmx [dot] com.

Encoding
--------
If you write Stata code containing non-ASCII characters in ST3, you may notice that these characters have been replaced with nonsense when you open the file in Stata's native do-file editor. This is because Stata and ST3 use different encodings. To avoid this issue, you can save your file with a different encoding in ST3. To do so, open File -> Save with encoding, and choose the appropriate encoding. Which encoding is right for you I would imagine depends on your Stata distribution. For most users in Europe and the U.S. you would probably get the correct result by choosing Western (Windows 1252). If that doesn't work, try saving with other encodings.

Acknowledgments
---------------
Thanks to Adrian Adermon and Daniel Forchheimer for helpful suggestions and to Sergio Correia for providing additional key bindings.
